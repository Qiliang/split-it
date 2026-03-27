import re
from dataclasses import dataclass
from typing import Optional


@dataclass
class MarkdownBlock:
    level: int
    title: str
    content: str
    parent: Optional['MarkdownBlock'] = None


    def to_markdown(self, include_parent: bool = False) -> str:
        parts: list[str] = []
        if include_parent:
            ancestors: list['MarkdownBlock'] = []
            node = self.parent
            while node:
                ancestors.append(node)
                node = node.parent
            for anc in reversed(ancestors):
                parts.append(f"{'#' * anc.level} {anc.title}")
                if anc.content:
                    parts.append(anc.content)
        parts.append(f"{'#' * self.level} {self.title}")
        if self.content:
            parts.append(self.content)
        return '\n\n'.join(parts)


_HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)')


def markdown_split_by_level(markdown_text: str, depth=3) -> list['MarkdownBlock']:
    """
    基于markdown的标题层级,将文档切分成多个子文档
    depth=4 那么，最顶层标题为1级标题，以下标题为2级标题，以此类推，直到depth级标题。
    在markdown中，# 为一级标题;## 为二级标题;以此类推，直到#### 为depth级标题。
    return:
        返回depth的MarkdownBlock。
        例如文档中有20个depth为4的标题，那么返回20个MarkdownBlock(parent中包含的MarkdownBlock为depth-1的标题,parent为depth-2的标题，以此类推)。
    """
    lines = markdown_text.split('\n')

    sections: list[tuple[int, int, str]] = []
    for i, line in enumerate(lines):
        m = _HEADING_RE.match(line)
        if m and len(m.group(1)) <= depth:
            sections.append((i, len(m.group(1)), m.group(2).strip()))

    ancestors: dict[int, MarkdownBlock] = {}
    result: list[MarkdownBlock] = []

    for idx, (line_no, level, title) in enumerate(sections):
        content_start = line_no + 1
        content_end = sections[idx + 1][0] if idx + 1 < len(sections) else len(lines)
        content = '\n'.join(lines[content_start:content_end]).strip()

        parent = None
        for lv in range(level - 1, 0, -1):
            if lv in ancestors:
                parent = ancestors[lv]
                break

        block = MarkdownBlock(level=level, title=title, content=content, parent=parent)

        ancestors[level] = block
        for lv in list(ancestors):
            if lv > level:
                del ancestors[lv]

        if level == depth:
            result.append(block)

    return result


if __name__ == "__main__":
    with open("markdown_doc/output.md", "r", encoding="utf-8") as f:
        md_text = f.read()
    blocks = markdown_split_by_level(md_text, depth=3)
    # print(blocks)
    for block in blocks:
        print(block.to_markdown())
        print("-"*100)