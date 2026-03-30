"""
Generate a formatted Word user manual from the Markdown source.

This keeps the editable source in Markdown while producing a presentable
`.docx` file for sharing with customers or internal teams.
"""
from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape


DEFAULT_SOURCE = Path(__file__).resolve().parent / "CAPACITY_OPTIMIZER_USER_MANUAL_CN.md"
DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / "Capacity_Optimizer_User_Manual_CN.docx"


@dataclass
class Block:
    kind: str
    text: str
    level: int = 0


def parse_markdown(source_text: str) -> list[Block]:
    blocks: list[Block] = []
    lines = source_text.splitlines()
    i = 0
    in_code = False
    code_lines: list[str] = []
    paragraph_lines: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph_lines
        if paragraph_lines:
            text = " ".join(line.strip() for line in paragraph_lines).strip()
            if text:
                blocks.append(Block("paragraph", text))
            paragraph_lines = []

    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if stripped.startswith("```"):
            flush_paragraph()
            if in_code:
                blocks.append(Block("code", "\n".join(code_lines)))
                code_lines = []
                in_code = False
            else:
                in_code = True
            i += 1
            continue

        if in_code:
            code_lines.append(line)
            i += 1
            continue

        if not stripped or stripped == "---":
            flush_paragraph()
            blocks.append(Block("blank", ""))
            i += 1
            continue

        heading_match = re.match(r"^(#{1,5})\s+(.*)$", stripped)
        if heading_match:
            flush_paragraph()
            blocks.append(Block("heading", heading_match.group(2).strip(), level=len(heading_match.group(1))))
            i += 1
            continue

        ordered_match = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if ordered_match:
            flush_paragraph()
            blocks.append(Block("ordered", ordered_match.group(2).strip()))
            i += 1
            continue

        bullet_match = re.match(r"^[-*]\s+(.*)$", stripped)
        if bullet_match:
            flush_paragraph()
            blocks.append(Block("bullet", bullet_match.group(1).strip()))
            i += 1
            continue

        paragraph_lines.append(line)
        i += 1

    flush_paragraph()
    if code_lines:
        blocks.append(Block("code", "\n".join(code_lines)))
    return blocks


def split_inline_code(text: str) -> list[tuple[str, bool]]:
    parts: list[tuple[str, bool]] = []
    current = []
    in_code = False
    for char in text:
        if char == "`":
            if current:
                parts.append(("".join(current), in_code))
                current = []
            in_code = not in_code
            continue
        current.append(char)
    if current:
        parts.append(("".join(current), in_code))
    return parts


def run_xml(text: str, *, code: bool = False, bold: bool = False) -> str:
    props = [
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>',
    ]
    if code:
        props = [
            '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Microsoft YaHei"/>',
            '<w:color w:val="1F1F1F"/>',
        ]
    if bold:
        props.append("<w:b/>")
    preserve = ' xml:space="preserve"' if text[:1].isspace() or text[-1:].isspace() or "  " in text else ""
    return f'<w:r><w:rPr>{"".join(props)}</w:rPr><w:t{preserve}>{escape(text)}</w:t></w:r>'


def inline_runs_xml(text: str) -> str:
    runs: list[str] = []
    for chunk, is_code in split_inline_code(text):
        if not chunk:
            continue
        runs.append(run_xml(chunk, code=is_code))
    return "".join(runs) or run_xml("")


def paragraph_xml(block: Block) -> str:
    if block.kind == "blank":
        return "<w:p/>"

    if block.kind == "heading":
        style = f"Heading{min(max(block.level, 1), 5)}"
        return (
            f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
            f"{inline_runs_xml(block.text)}</w:p>"
        )

    if block.kind == "bullet":
        return (
            "<w:p><w:pPr><w:pStyle w:val=\"ListBullet\"/>"
            "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr>"
            f"{inline_runs_xml(block.text)}</w:p>"
        )

    if block.kind == "ordered":
        return (
            "<w:p><w:pPr><w:pStyle w:val=\"ListNumber\"/>"
            "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
            f"{inline_runs_xml(block.text)}</w:p>"
        )

    if block.kind == "code":
        paragraphs = []
        for line in block.text.splitlines() or [""]:
            paragraphs.append(
                "<w:p><w:pPr><w:pStyle w:val=\"CodeBlock\"/></w:pPr>"
                f"{run_xml(line if line else ' ', code=True)}</w:p>"
            )
        return "".join(paragraphs)

    return f'<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>{inline_runs_xml(block.text)}</w:p>'


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="240" w:after="160"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>
      <w:b/>
      <w:color w:val="1F4E79"/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="200" w:after="140"/></w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="1F4E79"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="180" w:after="120"/></w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="2F5496"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="140" w:after="100"/></w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="385D8A"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading5">
    <w:name w:val="heading 5"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="4F81BD"/>
      <w:sz w:val="21"/>
      <w:szCs w:val="21"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720" w:hanging="360"/><w:spacing w:after="60"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListNumber">
    <w:name w:val="List Number"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720" w:hanging="360"/><w:spacing w:after="60"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="Code Block"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="360" w:right="120"/>
      <w:spacing w:after="40"/>
      <w:shd w:fill="F3F3F3"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Microsoft YaHei"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
</w:styles>
"""


def numbering_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>
      </w:rPr>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
"""


def document_xml(blocks: list[Block]) -> str:
    body = "".join(paragraph_xml(block) for block in blocks)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w10="urn:schemas-microsoft-com:office:word"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 mc:Ignorable="w14 wp14">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def document_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""


def content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def package_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def core_props_xml() -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Chemical Capacity Optimizer 使用手册</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>
"""


def app_props_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>RSCP</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"""


def build_docx(source_path: Path, output_path: Path) -> None:
    blocks = parse_markdown(source_path.read_text(encoding="utf-8"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", content_types_xml())
        package.writestr("_rels/.rels", package_rels_xml())
        package.writestr("word/document.xml", document_xml(blocks))
        package.writestr("word/styles.xml", styles_xml())
        package.writestr("word/numbering.xml", numbering_xml())
        package.writestr("word/_rels/document.xml.rels", document_rels_xml())
        package.writestr("docProps/core.xml", core_props_xml())
        package.writestr("docProps/app.xml", app_props_xml())


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate the formatted user manual DOCX from Markdown.")
    parser.add_argument("--source", default=str(DEFAULT_SOURCE), help="Markdown source path.")
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT), help="Output DOCX path.")
    args = parser.parse_args()

    build_docx(Path(args.source), Path(args.out))
    print(args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
