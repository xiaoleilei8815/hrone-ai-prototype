from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


CALLOUT_FILLS = {"E8EEF5", "E7F4F2", "FFF5D9", "FCEBEC"}
CODE_FILL = "F4F6F9"
IMAGE_NAMES = ("retention-calculation-flow.png", "retention-tool-architecture.png")


def escape_inline(text: str) -> str:
    text = text.replace("\\", "\\\\")
    for char in ("`", "*", "_", "["):
        text = text.replace(char, f"\\{char}")
    return text


def run_text(run_element) -> str:
    parts: list[str] = []
    for node in run_element.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag == qn("w:br"):
            parts.append("\n")
    return "".join(parts)


def formatted_run(run_element, *, escape: bool = True) -> str:
    text = run_text(run_element)
    if not text:
        return ""
    rendered = escape_inline(text) if escape else text
    properties = run_element.find(qn("w:rPr"))
    if properties is None:
        return rendered
    bold = properties.find(qn("w:b")) is not None
    italic = properties.find(qn("w:i")) is not None
    if bold and italic:
        return f"***{rendered}***"
    if bold:
        return f"**{rendered}**"
    if italic:
        return f"*{rendered}*"
    return rendered


def hyperlink_target(document, hyperlink_element) -> str | None:
    relationship_id = hyperlink_element.get(qn("r:id"))
    if not relationship_id:
        return None
    relationship = document.part.rels.get(relationship_id)
    return relationship.target_ref if relationship else None


def paragraph_markdown(document, paragraph: Paragraph, *, escape: bool = True) -> str:
    parts: list[str] = []
    for child in paragraph._p:
        if child.tag == qn("w:r"):
            parts.append(formatted_run(child, escape=escape))
        elif child.tag == qn("w:hyperlink"):
            label = "".join(formatted_run(run, escape=escape) for run in child.findall(qn("w:r")))
            target = hyperlink_target(document, child)
            parts.append(f"[{label}]({target})" if target else label)
    return "".join(parts).strip()


def paragraph_plain_text(paragraph: Paragraph) -> str:
    parts: list[str] = []
    for run in paragraph._p.iter(qn("w:r")):
        parts.append(run_text(run))
    return "".join(parts).strip()


def paragraph_fill(paragraph: Paragraph) -> str | None:
    properties = paragraph._p.find(qn("w:pPr"))
    if properties is None:
        return None
    shading = properties.find(qn("w:shd"))
    return shading.get(qn("w:fill")) if shading is not None else None


def numbering_format(document, paragraph: Paragraph) -> str | None:
    properties = paragraph._p.find(qn("w:pPr"))
    if properties is None:
        return None
    numbering = properties.find(qn("w:numPr"))
    if numbering is None:
        return None
    number_id_element = numbering.find(qn("w:numId"))
    if number_id_element is None:
        return None
    number_id = number_id_element.get(qn("w:val"))
    numbering_root = document.part.numbering_part.element
    num = next(
        (node for node in numbering_root.findall(qn("w:num")) if node.get(qn("w:numId")) == number_id),
        None,
    )
    if num is None:
        return None
    abstract_id = num.find(qn("w:abstractNumId"))
    if abstract_id is None:
        return None
    abstract_value = abstract_id.get(qn("w:val"))
    abstract = next(
        (
            node
            for node in numbering_root.findall(qn("w:abstractNum"))
            if node.get(qn("w:abstractNumId")) == abstract_value
        ),
        None,
    )
    if abstract is None:
        return None
    level = abstract.find(qn("w:lvl"))
    number_format = level.find(qn("w:numFmt")) if level is not None else None
    return number_format.get(qn("w:val")) if number_format is not None else None


def image_extension(content_type: str) -> str:
    return {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/svg+xml": ".svg",
    }.get(content_type, ".bin")


def extract_paragraph_images(document, paragraph: Paragraph, asset_dir: Path, image_index: int):
    outputs: list[tuple[str, int]] = []
    for blip in paragraph._p.iter(qn("a:blip")):
        relationship_id = blip.get(qn("r:embed"))
        if not relationship_id:
            continue
        part = document.part.related_parts[relationship_id]
        if image_index < len(IMAGE_NAMES):
            filename = IMAGE_NAMES[image_index]
        else:
            filename = f"document-image-{image_index + 1}{image_extension(part.content_type)}"
        asset_dir.mkdir(parents=True, exist_ok=True)
        (asset_dir / filename).write_bytes(part.blob)
        outputs.append((filename, image_index + 1))
        image_index += 1
    return outputs, image_index


def cell_text(cell) -> str:
    paragraphs = [paragraph_plain_text(paragraph) for paragraph in cell.paragraphs]
    value = "<br>".join(text for text in paragraphs if text)
    return value.replace("|", "\\|") or " "


def table_markdown(table: Table) -> list[str]:
    rows = [[cell_text(cell) for cell in row.cells] for row in table.rows]
    if not rows:
        return []
    column_count = max(len(row) for row in rows)
    normalized = [row + [" "] * (column_count - len(row)) for row in rows]
    output = ["| " + " | ".join(normalized[0]) + " |"]
    output.append("| " + " | ".join(["---"] * column_count) + " |")
    output.extend("| " + " | ".join(row) + " |" for row in normalized[1:])
    return output


def is_source_note(paragraph: Paragraph) -> bool:
    runs = list(paragraph._p.iter(qn("w:r")))
    if not runs:
        return False
    has_text = False
    for run in runs:
        if run_text(run):
            has_text = True
            properties = run.find(qn("w:rPr"))
            if properties is None or properties.find(qn("w:i")) is None:
                return False
    return has_text


def compress_blank_lines(lines: list[str]) -> str:
    text = "\n".join(lines).replace("\u00a0", " ")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def convert(input_path: Path, output_path: Path, asset_dir: Path) -> None:
    document = Document(input_path)
    output: list[str] = []
    image_index = 0

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, document._body)
            images, image_index = extract_paragraph_images(document, paragraph, asset_dir, image_index)
            if images:
                for filename, index in images:
                    label = "留存率计算与校验流程图" if index == 1 else "工具技术架构图"
                    relative_path = Path(asset_dir.name) / filename
                    output.extend([f"![{label}]({relative_path.as_posix()})", ""])
                continue

            plain = paragraph_plain_text(paragraph)
            if not plain:
                continue
            style = paragraph.style.name if paragraph.style is not None else "Normal"
            fill = paragraph_fill(paragraph)
            if fill == CODE_FILL:
                fence = "````" if "```" in plain else "```"
                output.extend([fence, plain, fence, ""])
                continue

            rendered = paragraph_markdown(document, paragraph)
            if fill in CALLOUT_FILLS:
                output.extend([f"> {rendered}", ""])
            elif style == "Doc Title":
                output.extend([f"# {' '.join(plain.split())}", ""])
            elif style == "Doc Subtitle":
                output.extend([f"> {plain}", ""])
            elif style == "Kicker":
                output.extend([f"**{plain}**", ""])
            elif style == "Small Meta":
                output.extend([f"- {rendered}", ""])
            elif style.startswith("Heading "):
                level = int(style.rsplit(" ", 1)[1])
                output.extend([f"{'#' * (level + 1)} {plain}", ""])
            elif numbering_format(document, paragraph) == "bullet":
                output.extend([f"- {rendered}", ""])
            elif numbering_format(document, paragraph) == "decimal":
                output.extend([f"1. {rendered}", ""])
            elif is_source_note(paragraph):
                output.extend([f"> *{escape_inline(plain)}*", ""])
            else:
                output.extend([rendered, ""])
        elif child.tag == qn("w:tbl"):
            output.extend(table_markdown(Table(child, document._body)))
            output.append("")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(compress_blank_lines(output), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert the retention handoff DOCX to GitHub Markdown.")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--assets", type=Path, required=True)
    args = parser.parse_args()
    convert(args.input.resolve(), args.output.resolve(), args.assets.resolve())


if __name__ == "__main__":
    main()
