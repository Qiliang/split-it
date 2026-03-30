import io
import os
import re
import struct
import zipfile
from dataclasses import dataclass
from typing import Optional
from xml.sax.saxutils import escape as _xml_escape
from langchain_core.documents.base import Document
from langchain_text_splitters import MarkdownTextSplitter,MarkdownHeaderTextSplitter


_HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)')


def markdown_split_by_level(markdown_text: str, depth=3) -> list[str]:
    """
    基于markdown的标题层级,将文档切分成多个子文档
    depth=4 那么，最顶层标题为1级标题，以下标题为2级标题，以此类推，直到depth级标题。
    在markdown中，# 为一级标题;## 为二级标题;以此类推，直到#### 为depth级标题。
    return:
        返回depth的MarkdownBlock。
        例如文档中有20个depth为4的标题，那么返回20个MarkdownBlock(parent中包含的MarkdownBlock为depth-1的标题,parent为depth-2的标题，以此类推)。
    """
    if depth == 1:
        splitter = MarkdownHeaderTextSplitter(headers_to_split_on=[("#","标题")])
    elif depth == 2:
        splitter = MarkdownHeaderTextSplitter(headers_to_split_on=[("#","一级标题"),("##","二级标题")])
    elif depth == 3:
        splitter = MarkdownHeaderTextSplitter(headers_to_split_on=[("#","一级标题"),("##","二级标题"),("###","三级标题")])
    elif depth == 4:
        splitter = MarkdownHeaderTextSplitter(headers_to_split_on=[("#","一级标题"),("##","二级标题"),("###","三级标题"),("####","四级标题")])
    else:
        raise ValueError(f"depth must be in [1,2,3,4], but got {depth}")
    blocks = splitter.split_text(markdown_text)
    return [str(block) for block in blocks if len(block.metadata.keys()) == depth]


def markdown_split_by_text(markdown_text: str, chunk_size=1000, chunk_overlap=100) -> list[str]:
    """
    基于markdown的文本，将文档切分成多个子文档
    return:
        返回MarkdownBlock。
    """
    splitter = MarkdownTextSplitter(chunk_size=chunk_size, chunk_overlap=chunk_overlap)
    return splitter.split_text(markdown_text)


# ── DOCX 常量 ────────────────────────────────────────────────────────────────

_IMAGE_RE = re.compile(r'!\[.*?\]\(\s*(.*?)\s*\)')

_MAX_WIDTH_EMU = 5943300  # 约 6.5 英寸，对应 A4 页面正文宽度

_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="word/document.xml"/>'
    '</Relationships>'
)

_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
    '<w:name w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>'
    '<w:rPr>'
    '<w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun" w:cs="SimSun"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    '</w:rPr>'
    '</w:style>'
    '</w:styles>'
)

_DOCUMENT_NS = (
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)

# ── 内部工具函数 ──────────────────────────────────────────────────────────────

def _get_image_dimensions(data: bytes) -> tuple[int, int]:
    """从 PNG 或 JPEG 二进制数据解析宽高（像素）。失败返回 (0, 0)。"""
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        w = struct.unpack('>I', data[16:20])[0]
        h = struct.unpack('>I', data[20:24])[0]
        return w, h
    if data[:2] == b'\xff\xd8':
        i = 2
        while i < len(data) - 8:
            if data[i] != 0xff:
                break
            marker = data[i + 1]
            if marker in (0xc0, 0xc1, 0xc2):
                h = struct.unpack('>H', data[i + 5:i + 7])[0]
                w = struct.unpack('>H', data[i + 7:i + 9])[0]
                return w, h
            seg_len = struct.unpack('>H', data[i + 2:i + 4])[0]
            i += 2 + seg_len
    return 0, 0


def _calc_emu(data: bytes, dpi: int = 96) -> tuple[int, int]:
    """计算图片的 EMU 尺寸，超出页面宽度时等比缩放。"""
    w_px, h_px = _get_image_dimensions(data)
    if w_px == 0 or h_px == 0:
        return _MAX_WIDTH_EMU, int(_MAX_WIDTH_EMU * 9 / 16)
    w_emu = int(w_px * 914400 / dpi)
    h_emu = int(h_px * 914400 / dpi)
    if w_emu > _MAX_WIDTH_EMU:
        h_emu = int(h_emu * _MAX_WIDTH_EMU / w_emu)
        w_emu = _MAX_WIDTH_EMU
    return w_emu, h_emu


def _run(text: str, bold: bool = False) -> str:
    """生成一个 <w:r> 元素。"""
    t = _xml_escape(text)
    preserve = ' xml:space="preserve"' if text != text.strip() else ''
    rpr = '<w:rPr><w:b/><w:bCs/></w:rPr>' if bold else ''
    return f'<w:r>{rpr}<w:t{preserve}>{t}</w:t></w:r>'


def _para(runs_xml: str) -> str:
    """生成一个 <w:p> 元素。"""
    return f'<w:p>{runs_xml}</w:p>'


def _text_para(full_text: str, bold_prefix: str = '') -> str:
    """生成带可选粗体前缀的文本段落。"""
    if bold_prefix and full_text.startswith(bold_prefix):
        runs = _run(bold_prefix, bold=True) + _run(full_text[len(bold_prefix):])
    else:
        runs = _run(full_text)
    return _para(runs)


def _separator_para() -> str:
    """生成一个带下边框的水平分隔线段落。"""
    return (
        '<w:p>'
        '<w:pPr>'
        '<w:pBdr>'
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
        '</w:pBdr>'
        '<w:spacing w:before="120" w:after="120"/>'
        '</w:pPr>'
        '</w:p>'
    )


def _image_para(r_id: str, cx: int, cy: int, draw_id: int, name: str) -> str:
    """生成包含内联图片的段落。"""
    esc_name = _xml_escape(name)
    return (
        '<w:p><w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{draw_id}" name="{esc_name}"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic>'
        '<pic:nvPicPr>'
        f'<pic:cNvPr id="0" name="{esc_name}"/>'
        '<pic:cNvPicPr/>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        f'<a:blip r:embed="{r_id}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing></w:r></w:p>'
    )


def _parse_qa_pairs(text: str) -> list[tuple[str, str]]:
    """
    从 Markdown 文本中解析 QA 对列表。
    格式：
      - Q: 问题
        A: 答案（可包含图片引用）
    """
    pairs: list[tuple[str, str]] = []
    cur_q: str | None = None
    cur_a_parts: list[str] = []

    for line in text.split('\n'):
        q_m = re.match(r'^- Q:\s*(.*)', line)
        a_m = re.match(r'^\s+A:\s*(.*)', line)
        if q_m:
            if cur_q is not None:
                pairs.append((cur_q, ' '.join(cur_a_parts)))
            cur_q = q_m.group(1).strip()
            cur_a_parts = []
        elif a_m and cur_q is not None:
            cur_a_parts.append(a_m.group(1).strip())
        elif cur_q is not None and cur_a_parts and line.strip():
            cur_a_parts.append(line.strip())

    if cur_q is not None:
        pairs.append((cur_q, ' '.join(cur_a_parts)))
    return pairs


# ── 公开函数 ──────────────────────────────────────────────────────────────────

def markdown_to_docx(markdown_text: str, image_dir: str = '.') -> bytes:
    """
    将简单 QA Markdown 格式转换为 DOCX 字节流，不依赖任何第三方库。

    Markdown 格式示例：
      - Q: 问题文本
        A: 答案文本 ![](_images/image_1.png)

    Args:
        markdown_text: Markdown 字符串。
        image_dir: 图片文件的根目录，用于解析相对路径（默认为当前目录）。

    Returns:
        DOCX 文件的字节内容，可直接写入 .docx 文件。
    """
    qa_pairs = _parse_qa_pairs(markdown_text)

    # 收集所有图片：img_path -> (rId, data, ext, media_name)
    images: dict[str, tuple[str, bytes, str, str]] = {}
    rId_num = 1

    def _add_image(img_path: str) -> None:
        nonlocal rId_num
        if img_path in images:
            return
        full = os.path.join(image_dir, img_path)
        if not os.path.exists(full):
            return
        with open(full, 'rb') as fh:
            data = fh.read()
        ext = os.path.splitext(img_path)[1].lstrip('.').lower() or 'png'
        images[img_path] = (f'rId{rId_num}', data, ext, f'image{rId_num}.{ext}')
        rId_num += 1

    for _, a_text in qa_pairs:
        for m in _IMAGE_RE.finditer(a_text):
            _add_image(m.group(1).strip())

    # 构建段落 XML
    body_parts: list[str] = []
    draw_id = 1

    for q_text, a_text in qa_pairs:
        q_clean = _IMAGE_RE.sub('', q_text).strip()
        body_parts.append(_text_para(f'Q: {q_clean}', bold_prefix='Q: '))

        a_clean = _IMAGE_RE.sub('', a_text).strip()
        body_parts.append(_text_para(f'A: {a_clean}', bold_prefix='A: '))

        for m in _IMAGE_RE.finditer(a_text):
            info = images.get(m.group(1).strip())
            if info:
                r_id, data, _, media_name = info
                cx, cy = _calc_emu(data)
                body_parts.append(_image_para(r_id, cx, cy, draw_id, media_name))
                draw_id += 1

        body_parts.append(_separator_para())

    body_xml = ''.join(body_parts)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {_DOCUMENT_NS}>'
        '<w:body>'
        f'{body_xml}'
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
        '</w:body>'
        '</w:document>'
    )

    # 构建 word/_rels/document.xml.rels
    rel_items = ''.join(
        f'<Relationship Id="{r_id}"'
        f' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
        f' Target="media/{media_name}"/>'
        for r_id, _, _, media_name in images.values()
    )
    doc_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{rel_items}'
        '</Relationships>'
    )

    # 构建 [Content_Types].xml
    ext_types: dict[str, str] = {}
    for _, _, ext, _ in images.values():
        if ext not in ext_types:
            ct = 'image/jpeg' if ext in ('jpg', 'jpeg') else f'image/{ext}'
            ext_types[ext] = ct
    img_defaults = ''.join(
        f'<Default Extension="{ext}" ContentType="{ct}"/>'
        for ext, ct in ext_types.items()
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels"'
        ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument'
        '.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument'
        '.wordprocessingml.styles+xml"/>'
        f'{img_defaults}'
        '</Types>'
    )

    # 打包成 ZIP（即 DOCX）
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml)
        zf.writestr('_rels/.rels', _ROOT_RELS)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/_rels/document.xml.rels', doc_rels_xml)
        zf.writestr('word/styles.xml', _STYLES_XML)
        for _, (_, data, _, media_name) in images.items():
            zf.writestr(f'word/media/{media_name}', data)

    return buf.getvalue()




if __name__ == "__main__":
    with open("markdown_doc/_output.md", "r", encoding="utf-8") as f:
        md_text = f.read()
    # blocks = markdown_split_by_level(md_text, depth=3)
    blocks = markdown_split_by_text(md_text, chunk_size=1000, chunk_overlap=200)
    # print(blocks)
    for block in blocks:
        print(block)
        # print(block.to_markdown())
        print("-"*100)
