import json
import os
from pathlib import Path
import re

from tqdm import tqdm
import mammoth
import markdownify
import openai
from paser import markdown_split_by_level


from markitdown.converter_utils.docx.pre_process import pre_process_docx


def docx_to_markdown(docx_file_path: str, output_dir: str | None = None) -> str:
    """将 DOCX 文件转换为 Markdown，提取图片并以相对路径引用。"""
    docx_path = Path(docx_file_path).resolve()
    if output_dir is None:
        output_dir = str(docx_path.parent)

    images_dir_name = "_images"
    images_dir = os.path.join(output_dir, images_dir_name)

    image_counter = 0

    def handle_image(image):
        nonlocal image_counter
        image_counter += 1
        ext = image.content_type.split("/")[-1]
        if ext == "jpeg":
            ext = "jpg"
        filename = f"image_{image_counter}.{ext}"
        filepath = os.path.join(images_dir, filename)

        os.makedirs(images_dir, exist_ok=True)
        with image.open() as img_stream:
            with open(filepath, "wb") as f:
                f.write(img_stream.read())

        return {"src": f"{images_dir_name}/{filename}"}

    with open(docx_file_path, "rb") as f:
        pre_processed = pre_process_docx(f)

    html = mammoth.convert_to_html(
        pre_processed,
        convert_image=mammoth.images.img_element(handle_image),
    ).value

    md_text = markdownify.markdownify(html, heading_style="ATX")
    return md_text.strip()


def _llm(content: str, api_key: str = "", base_url: str = "") -> str:
    client = openai.OpenAI(api_key=api_key, base_url=base_url)
    completion = client.chat.completions.create(
        model="qwen-plus",
        messages=[{"role": "user", "content": content}],
    )
    return completion.choices[0].message.content


def extract_qa_pairs(md_file_path: str,depth:int=3) -> list[tuple[str, str]]:
    """从 Markdown 文本中提取有价值的问答对（QA Pairs）。"""
    with open(md_file_path, "r", encoding="utf-8") as f:
        md_text = f.read()
    blocks = markdown_split_by_level(md_text, depth=depth)
    with open("QAprompt.txt", "r", encoding="utf-8") as f:
        prompt_template = f.read()
    
    for i, block in tqdm(enumerate(blocks)):
        print(block.to_markdown())
        prompt = prompt_template.replace("$DOC_CONTENT$", block.to_markdown(include_parent=True))
        response_text = _llm(prompt)
        # 通过正则提取```markdown 和 ``` 之间的内容
        response_text = re.search(r"```markdown(.*?)\n```", response_text, re.DOTALL).group(1)
        with open(f"markdown_doc/qa_pairs_{i}.md", "w", encoding="utf-8") as f:
            f.write(response_text)


def main():
    md_text = docx_to_markdown("/Users/xiaoql/Downloads/发票管理使用手册-终版.docx")
    with open("output.md", "w", encoding="utf-8") as f:
        f.write(md_text)
    print("转换完成，输出文件: output.md")


if __name__ == "__main__":
    # main()
    # extract_qa_pairs("markdown_doc/output.md")
    with open("markdown_doc/qa_pairs_1.md", "r", encoding="utf-8") as f:
        response_text = f.read()
        response_text = re.search(r"```markdown(.*?)\n```", response_text, re.DOTALL).group(1)
    print(response_text)