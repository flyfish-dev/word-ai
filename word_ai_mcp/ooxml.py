from __future__ import annotations

import copy
import csv
import csv
import difflib
import hashlib
import json
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from lxml import etree

from .types import Anchor, ValidationIssue, ValidationReport

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W_NS, "r": R_NS, "wp": WP_NS, "w14": W14_NS}
REL_NS = {"rel": PKG_REL_NS}
CT = {"ct": CT_NS}

DOCUMENT_XML = "word/document.xml"
DOCUMENT_RELS = "word/_rels/document.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
COMMENTS_XML = "word/comments.xml"
STYLES_XML = "word/styles.xml"
NUMBERING_XML = "word/numbering.xml"
SETTINGS_XML = "word/settings.xml"

COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

COMPLEX_XPATH = ".//w:tbl | .//w:drawing | .//w:pict | .//w:fldSimple | .//w:instrText | .//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference | .//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"
HIGH_RISK_OPS = {
    "replace_paragraph_text",
    "insert_paragraph_after",
    "insert_paragraph_before",
    "replace_table_cell_text",
    "append_table_row",
    "wrap_paragraph_with_content_control",
    "add_comment",
}


def qn(tag: str) -> str:
    prefix, name = tag.split(":", 1)
    ns = {"w": W_NS, "r": R_NS, "wp": WP_NS, "w14": W14_NS}[prefix]
    return f"{{{ns}}}{name}"


def _parse_xml(data: bytes) -> etree._ElementTree:
    parser = etree.XMLParser(remove_blank_text=False, recover=False, huge_tree=True)
    return etree.fromstring(data, parser=parser).getroottree()


def _serialize_xml(tree: etree._ElementTree | etree._Element) -> bytes:
    root = tree.getroot() if isinstance(tree, etree._ElementTree) else tree
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _canonical_hash(el: etree._Element) -> str:
    try:
        data = etree.tostring(el, method="c14n", with_comments=True)
    except Exception:
        data = etree.tostring(el, encoding="UTF-8")
    return hashlib.sha256(data).hexdigest()


def _sha(data: bytes | str) -> str:
    if isinstance(data, str):
        data = data.encode("utf-8")
    return hashlib.sha256(data).hexdigest()


def _zip_read(path: str | Path, member: str) -> bytes:
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read(member)


def _zip_has(path: str | Path, member: str) -> bool:
    with zipfile.ZipFile(path, "r") as zf:
        return member in zf.namelist()


def _zip_hashes(path: str | Path) -> dict[str, str]:
    with zipfile.ZipFile(path, "r") as zf:
        return {i.filename: hashlib.sha256(zf.read(i.filename)).hexdigest() for i in zf.infolist() if not i.is_dir()}


def package_manifest(docx_path: str | Path, include_hashes: bool = True) -> dict[str, Any]:
    """Return DOCX package entries. Read-only and useful for Codex sanity checks."""
    docx_path = Path(docx_path)
    entries: list[dict[str, Any]] = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            item: dict[str, Any] = {
                "name": info.filename,
                "compress_type": info.compress_type,
                "file_size": info.file_size,
                "compress_size": info.compress_size,
                "date_time": info.date_time,
            }
            if include_hashes:
                item["sha256"] = _sha(zf.read(info.filename))
            entries.append(item)
    return {
        "path": str(docx_path),
        "sha256": _sha(docx_path.read_bytes()),
        "size_bytes": docx_path.stat().st_size,
        "part_count": len(entries),
        "parts": entries,
    }


def atomic_copy_docx(src: str | Path, dst: str | Path) -> None:
    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    tmp = Path(dst).with_suffix(Path(dst).suffix + ".tmp")
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)


def rewrite_zip_member(src_docx: str | Path, dst_docx: str | Path, replacements: dict[str, bytes]) -> None:
    """Rewrite selected members while byte-preserving all other ZIP entries."""
    src_docx = Path(src_docx)
    dst_docx = Path(dst_docx)
    dst_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=str(dst_docx.parent)) as tmp:
        tmp_name = tmp.name
    try:
        with zipfile.ZipFile(src_docx, "r") as zin, zipfile.ZipFile(tmp_name, "w") as zout:
            seen: set[str] = set()
            for item in zin.infolist():
                data = replacements.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
                seen.add(item.filename)
            for name, data in replacements.items():
                if name not in seen:
                    zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tmp_name, dst_docx)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def load_document_tree(docx_path: str | Path) -> etree._ElementTree:
    return _parse_xml(_zip_read(docx_path, DOCUMENT_XML))


def paragraph_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NS))


def element_text(el: etree._Element) -> str:
    return "\n".join(filter(None, (paragraph_text(p).strip() for p in el.xpath(".//w:p", namespaces=NS))))


def _direct_paragraphs(container: etree._Element) -> list[etree._Element]:
    return [c for c in container if c.tag == qn("w:p")]


def sdt_tag(sdt: etree._Element) -> str | None:
    tag = sdt.find("./w:sdtPr/w:tag", namespaces=NS)
    if tag is not None:
        return tag.get(qn("w:val"))
    return None


def sdt_alias(sdt: etree._Element) -> str | None:
    alias = sdt.find("./w:sdtPr/w:alias", namespaces=NS)
    if alias is not None:
        return alias.get(qn("w:val"))
    return None


def sdt_id(sdt: etree._Element) -> str | None:
    sid = sdt.find("./w:sdtPr/w:id", namespaces=NS)
    if sid is not None:
        return sid.get(qn("w:val"))
    return None


def _style_id_for_paragraph(p: etree._Element) -> str | None:
    pstyle = p.find("./w:pPr/w:pStyle", namespaces=NS)
    return pstyle.get(qn("w:val")) if pstyle is not None else None


def _style_name_for_id(styles: dict[str, dict[str, Any]], style_id: str | None) -> str | None:
    if not style_id:
        return None
    info = styles.get(style_id)
    return str(info.get("name")) if info and info.get("name") is not None else None


def _outline_level_value(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        raw = int(value)
    except (TypeError, ValueError):
        return None
    if 0 <= raw <= 8:
        return raw + 1
    return None


def _heading_level_from_name(value: str | None) -> int | None:
    if not value:
        return None
    normalized = re.sub(r"\s+", "", value.strip())
    match = re.match(r"(?i)^heading([1-9])$", normalized) or re.match(r"^标题([1-9])$", normalized)
    if match:
        return int(match.group(1))
    return None


def _is_toc_field_instruction(instr: str | None) -> bool:
    if not instr:
        return False
    return bool(re.match(r"^\s*TOC(?:\s|\\|$)", instr, re.I))


def _has_toc_field_instruction(p: etree._Element) -> bool:
    for fld in p.xpath(".//w:fldSimple", namespaces=NS):
        if _is_toc_field_instruction(fld.get(qn("w:instr"))):
            return True
    for instr in p.xpath(".//w:instrText", namespaces=NS):
        if _is_toc_field_instruction(instr.text or ""):
            return True
    return False


def _has_toc_reference_field(p: etree._Element) -> bool:
    values = [fld.get(qn("w:instr")) or "" for fld in p.xpath(".//w:fldSimple", namespaces=NS)]
    values.extend(instr.text or "" for instr in p.xpath(".//w:instrText", namespaces=NS))
    return any(re.search(r"\b(?:PAGEREF|HYPERLINK)\b.*_Toc", value, re.I) for value in values)


def _is_toc_style(style_id: str | None, styles: dict[str, dict[str, Any]] | None = None) -> bool:
    values = [style_id or ""]
    if styles and style_id and style_id in styles:
        values.append(str(styles[style_id].get("name") or ""))
    for value in values:
        normalized = re.sub(r"\s+", "", value.strip()).lower()
        if not normalized:
            continue
        if re.match(r"^toc\d*$", normalized) or normalized in {"tocheading", "tableofcontents"}:
            return True
        if normalized.startswith("toc") and ("heading" in normalized or normalized[3:].isdigit()):
            return True
        if normalized.startswith("目录") or normalized.startswith("wpsoffice手动目录"):
            return True
        if normalized.startswith("tableofcontents"):
            return True
    return False


def _paragraph_direct_outline_level(p: etree._Element) -> int | None:
    outline = p.find("./w:pPr/w:outlineLvl", namespaces=NS)
    return _outline_level_value(outline.get(qn("w:val")) if outline is not None else None)


def _heading_level(style_id: str | None, styles: dict[str, dict[str, Any]] | None = None, paragraph: etree._Element | None = None) -> int | None:
    if _is_toc_style(style_id, styles):
        return None
    if paragraph is not None:
        direct = _paragraph_direct_outline_level(paragraph)
        if direct is not None:
            return direct
    direct_style = _heading_level_from_name(style_id)
    if direct_style is not None:
        return direct_style
    if styles and style_id:
        info = styles.get(style_id)
        if info:
            if _is_toc_style(style_id, styles):
                return None
            name_level = _heading_level_from_name(str(info.get("name") or ""))
            if name_level is not None:
                return name_level
            outline_level = _outline_level_value(str(info.get("outline_level")) if info.get("outline_level") is not None else None)
            if outline_level is not None:
                return outline_level
    return None


def _load_paragraph_styles(docx_path: str | Path) -> dict[str, dict[str, Any]]:
    if not _zip_has(docx_path, STYLES_XML):
        return {}
    tree = _parse_xml(_zip_read(docx_path, STYLES_XML))
    styles: dict[str, dict[str, Any]] = {}
    for style in tree.getroot().xpath("//w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("w:styleId"))
        if not style_id:
            continue
        name = style.find("./w:name", namespaces=NS)
        based_on = style.find("./w:basedOn", namespaces=NS)
        outline = style.find("./w:pPr/w:outlineLvl", namespaces=NS)
        styles[style_id] = {
            "style_id": style_id,
            "name": name.get(qn("w:val")) if name is not None else None,
            "based_on": based_on.get(qn("w:val")) if based_on is not None else None,
            "outline_level": outline.get(qn("w:val")) if outline is not None else None,
            "is_toc": False,
        }
    for style_id, info in styles.items():
        info["is_toc"] = _is_toc_style(style_id, styles)
        info["heading_level"] = _heading_level(style_id, styles)
    return styles


def _is_toc_sdt_descendant(el: etree._Element) -> bool:
    cur: etree._Element | None = el
    while cur is not None:
        if cur.tag == qn("w:sdt"):
            values: list[str] = []
            for node in cur.xpath("./w:sdtPr//w:alias | ./w:sdtPr//w:tag | ./w:sdtPr//w:docPartGallery", namespaces=NS):
                value = node.get(qn("w:val"))
                if value:
                    values.append(value)
            raw = " ".join(values).lower()
            normalized = re.sub(r"\s+", "", raw)
            if "tableofcontents" in normalized or "目录" in raw or re.search(r"(^|[^a-z])toc([^a-z]|$)", raw):
                return True
        cur = cur.getparent()
    return False


def _toc_paragraph_indices(root: etree._Element, styles: dict[str, dict[str, Any]]) -> set[int]:
    paragraphs = root.xpath("//w:body//w:p", namespaces=NS)
    indices: set[int] = set()
    field_stack: list[dict[str, Any]] = []

    def active_toc_field() -> bool:
        return any(bool(ctx.get("is_toc")) for ctx in field_stack)

    for idx, p in enumerate(paragraphs, start=1):
        style_id = _style_id_for_paragraph(p)
        explicit_toc = _is_toc_style(style_id, styles) or _is_toc_sdt_descendant(p) or _has_toc_field_instruction(p) or _has_toc_reference_field(p)
        if active_toc_field() and not explicit_toc and paragraph_text(p).strip():
            # Some DOCX producers leave the TOC complex field unclosed. Do not let a
            # leaked TOC field consume the body outline after the visible TOC block.
            field_stack = [ctx for ctx in field_stack if not ctx.get("is_toc")]
        is_toc = active_toc_field() or explicit_toc
        for fld in p.xpath(".//w:fldSimple", namespaces=NS):
            if _is_toc_field_instruction(fld.get(qn("w:instr"))):
                is_toc = True
        for el in p.iter():
            if not isinstance(el.tag, str):
                continue
            if el.tag == qn("w:fldChar"):
                kind = el.get(qn("w:fldCharType"))
                if kind == "begin":
                    field_stack.append({"instr": "", "is_toc": False})
                elif kind == "separate":
                    if active_toc_field():
                        is_toc = True
                elif kind == "end":
                    if active_toc_field():
                        is_toc = True
                    if field_stack:
                        field_stack.pop()
            elif el.tag == qn("w:instrText") and field_stack:
                field_stack[-1]["instr"] += el.text or ""
                if _is_toc_field_instruction(field_stack[-1]["instr"]):
                    field_stack[-1]["is_toc"] = True
                    is_toc = True
        if active_toc_field():
            is_toc = True
        if is_toc:
            indices.add(idx)
    return indices


def _element_path(el: etree._Element) -> str:
    parts: list[str] = []
    cur: etree._Element | None = el
    while cur is not None and isinstance(cur.tag, str):
        parent = cur.getparent()
        tag = etree.QName(cur).localname
        if parent is not None:
            same = [c for c in parent if c.tag == cur.tag]
            idx = same.index(cur) + 1 if cur in same else 1
            parts.append(f"{tag}[{idx}]")
        else:
            parts.append(tag)
        cur = parent
    return "/" + "/".join(reversed(parts))


def _ancestor_sdt_tag(el: etree._Element) -> str | None:
    cur: etree._Element | None = el
    while cur is not None:
        if cur.tag == qn("w:sdt"):
            return sdt_tag(cur)
        cur = cur.getparent()
    return None


def _ancestor_table_index(el: etree._Element, root: etree._Element) -> int | None:
    cur: etree._Element | None = el
    table: etree._Element | None = None
    while cur is not None:
        if cur.tag == qn("w:tbl"):
            table = cur
            break
        cur = cur.getparent()
    if table is None:
        return None
    tables = root.xpath("//w:body//w:tbl", namespaces=NS)
    try:
        return tables.index(table) + 1
    except ValueError:
        return None


def _has_complex_content(el: etree._Element) -> bool:
    return bool(el.xpath(COMPLEX_XPATH, namespaces=NS))


def _complex_summary(el: etree._Element) -> dict[str, int]:
    return {
        "tables": len(el.xpath(".//w:tbl", namespaces=NS)),
        "drawings": len(el.xpath(".//w:drawing | .//w:pict", namespaces=NS)),
        "fields": len(el.xpath(".//w:fldSimple | .//w:instrText", namespaces=NS)),
        "comments": len(el.xpath(".//w:commentRangeStart | .//w:commentReference", namespaces=NS)),
        "tracked_changes": len(el.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo", namespaces=NS)),
    }


def _assert_expected_text(actual: str, expected_text: str | None = None, expected_sha256: str | None = None, label: str = "target") -> None:
    if expected_text is not None and actual != expected_text:
        raise ValueError(f"{label}: expected_text does not match current text")
    if expected_sha256 is not None and _sha(actual) != expected_sha256:
        raise ValueError(f"{label}: expected_sha256 does not match current text")


def inspect_docx(docx_path: str | Path, include_text: bool = False, max_preview: int = 240) -> dict[str, Any]:
    docx_path = Path(docx_path)
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    paragraphs = root.xpath("//w:body//w:p", namespaces=NS)
    tables = root.xpath("//w:body//w:tbl", namespaces=NS)
    images = root.xpath("//w:drawing | //w:pict", namespaces=NS)
    fields = root.xpath("//w:fldSimple | //w:instrText", namespaces=NS)
    tracked = root.xpath("//w:ins | //w:del | //w:moveFrom | //w:moveTo", namespaces=NS)
    comment_refs = root.xpath("//w:commentRangeStart | //w:commentReference", namespaces=NS)

    anchors = list_anchors(docx_path, max_preview=max_preview)
    text = "\n".join(paragraph_text(p) for p in paragraphs) if include_text else None
    parts: list[str] = []
    comments = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        parts = [i.filename for i in zf.infolist() if not i.is_dir()]
        if COMMENTS_XML in parts:
            try:
                ctree = _parse_xml(zf.read(COMMENTS_XML))
                comments = ctree.getroot().xpath("//w:comment", namespaces=NS)
            except Exception:
                comments = []
    return {
        "path": str(docx_path),
        "sha256": hashlib.sha256(docx_path.read_bytes()).hexdigest(),
        "size_bytes": docx_path.stat().st_size,
        "parts_count": len(parts),
        "paragraph_count": len(paragraphs),
        "table_count": len(tables),
        "image_count": len(images),
        "field_count": len(fields),
        "comment_count": len(comments),
        "comment_reference_count": len(comment_refs),
        "tracked_change_count": len(tracked),
        "content_control_count": sum(1 for a in anchors if a.kind == "content_control"),
        "heading_count": sum(1 for a in anchors if a.kind == "heading"),
        "bookmark_count": sum(1 for a in anchors if a.kind == "bookmark"),
        "anchors": [a.to_dict() for a in anchors],
        "text": text,
    }


def health_check(docx_path: str | Path, max_items: int = 50) -> dict[str, Any]:
    """Read-only report tuned for safe AI editing decisions."""
    info = inspect_docx(docx_path)
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    sdts = root.xpath("//w:sdt", namespaces=NS)
    tags: list[str] = [sdt_tag(s) or "" for s in sdts]
    tag_counts: dict[str, int] = {}
    for tag in tags:
        if tag:
            tag_counts[tag] = tag_counts.get(tag, 0) + 1
    duplicate_tags = sorted(k for k, v in tag_counts.items() if v > 1)
    complex_controls = []
    for s in sdts:
        tag = sdt_tag(s)
        summary = _complex_summary(s)
        if any(summary.values()):
            complex_controls.append({"tag": tag, "alias": sdt_alias(s), "complexity": summary, "path": _element_path(s)})
    paras = root.xpath("//w:body//w:p", namespaces=NS)
    para_ids = [p.get(qn("w14:paraId")) for p in paras if p.get(qn("w14:paraId"))]
    duplicate_para_ids = sorted({x for x in para_ids if para_ids.count(x) > 1})
    tables = list_tables(docx_path)
    risks: list[dict[str, str]] = []
    if info["content_control_count"] == 0:
        risks.append({"severity": "warning", "code": "no_content_controls", "message": "No content controls found. Prefer anchoring template sections before AI editing."})
    if duplicate_tags:
        risks.append({"severity": "error", "code": "duplicate_content_control_tags", "message": f"Duplicate content-control tags: {duplicate_tags[:max_items]}"})
    if complex_controls:
        risks.append({"severity": "warning", "code": "complex_content_controls", "message": "Some content controls contain tables/images/fields/comments/tracked changes; text replacement should be blocked unless explicitly allowed."})
    if info["field_count"]:
        risks.append({"severity": "info", "code": "fields_present", "message": "Fields are present. Avoid editing field code; refresh fields in Word after final edits."})
    if info["tracked_change_count"]:
        risks.append({"severity": "warning", "code": "tracked_changes_present", "message": "Tracked changes are present. Avoid direct replacements inside revisions unless reviewed."})
    return {
        "ok_for_ai_editing": not any(r["severity"] == "error" for r in risks),
        "metrics": {k: v for k, v in info.items() if k.endswith("_count") or k in ["parts_count", "size_bytes", "sha256"]},
        "duplicate_content_control_tags": duplicate_tags,
        "duplicate_para_ids": duplicate_para_ids,
        "complex_content_controls": complex_controls[:max_items],
        "tables": tables["tables"][:max_items],
        "risks": risks,
        "recommended_policy": {
            "prefer_content_control_ops": True,
            "require_expected_old_text_or_sha_for_writes": True,
            "default_allow_complex_content": False,
            "default_write_mode": "new_file_with_audit",
        },
    }


def list_anchors(docx_path: str | Path, max_preview: int = 240) -> list[Anchor]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    styles = _load_paragraph_styles(docx_path)
    toc_paragraphs = _toc_paragraph_indices(root, styles)
    anchors: list[Anchor] = []

    for sdt in root.xpath("//w:sdt", namespaces=NS):
        tag = sdt_tag(sdt)
        alias = sdt_alias(sdt)
        sid = sdt_id(sdt)
        content = element_text(sdt)
        preview = content[:max_preview]
        label = alias or tag or sid or "content_control"
        anchors.append(
            Anchor(
                anchor_id=f"cc:{tag or sid or hashlib.sha1(_element_path(sdt).encode()).hexdigest()[:12]}",
                kind="content_control",
                label=label,
                path=_element_path(sdt),
                text_preview=preview,
                extra={"tag": tag, "alias": alias, "id": sid, "text_sha256": _sha(content), "complexity": _complex_summary(sdt)},
            )
        )

    outline_stack: list[str] = []
    para_idx = 0
    for p in root.xpath("//w:body//w:p", namespaces=NS):
        para_idx += 1
        style_id = _style_id_for_paragraph(p)
        is_toc = para_idx in toc_paragraphs
        level = None if is_toc else _heading_level(style_id, styles, paragraph=p)
        txt = paragraph_text(p).strip()
        if level and txt:
            while len(outline_stack) >= level:
                outline_stack.pop()
            outline_stack.append(txt)
            anchors.append(
                Anchor(
                    anchor_id=f"heading:{para_idx}",
                    kind="heading",
                    label=" > ".join(outline_stack),
                    path=_element_path(p),
                    text_preview=txt[:max_preview],
                    style_id=style_id,
                    level=level,
                    extra={"paragraph_index": para_idx, "paraId": p.get(qn("w14:paraId")), "style_name": _style_name_for_id(styles, style_id), "is_toc": False, "text_sha256": _sha(txt)},
                )
            )
        elif txt and not is_toc:
            para_id = p.get(qn("w14:paraId"))
            if para_id:
                anchors.append(
                    Anchor(
                        anchor_id=f"p:{para_id}",
                        kind="paragraph",
                        label=f"Paragraph {para_idx}",
                        path=_element_path(p),
                        text_preview=txt[:max_preview],
                        style_id=style_id,
                        extra={"paragraph_index": para_idx, "paraId": para_id, "style_name": _style_name_for_id(styles, style_id), "content_control_tag": _ancestor_sdt_tag(p), "is_toc": False, "text_sha256": _sha(txt)},
                    )
                )

    for bm in root.xpath("//w:bookmarkStart", namespaces=NS):
        name = bm.get(qn("w:name"))
        if name and not name.startswith("_"):
            anchors.append(
                Anchor(
                    anchor_id=f"bookmark:{name}",
                    kind="bookmark",
                    label=name,
                    path=_element_path(bm),
                    text_preview="",
                    extra={"name": name, "id": bm.get(qn("w:id"))},
                )
            )
    return anchors


def get_outline(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    paras = root.xpath("//w:body//w:p", namespaces=NS)
    styles = _load_paragraph_styles(docx_path)
    toc_paragraphs = _toc_paragraph_indices(root, styles)
    headings: list[dict[str, Any]] = []
    stack: list[dict[str, Any]] = []
    for i, p in enumerate(paras, start=1):
        if i in toc_paragraphs:
            continue
        style_id = _style_id_for_paragraph(p)
        level = _heading_level(style_id, styles, paragraph=p)
        text = paragraph_text(p).strip()
        if not level or not text:
            continue
        while stack and int(stack[-1]["level"]) >= level:
            ended = stack.pop()
            for h in reversed(headings):
                if h["paragraph_index"] == ended["paragraph_index"]:
                    h["end_paragraph_index"] = i - 1
                    break
        entry = {
            "anchor_id": f"heading:{i}",
            "paragraph_index": i,
            "paraId": p.get(qn("w14:paraId")),
            "level": level,
            "style_id": style_id,
            "style_name": _style_name_for_id(styles, style_id),
            "is_toc": False,
            "text": text,
            "path": _element_path(p),
            "end_paragraph_index": len(paras),
        }
        headings.append(entry)
        stack.append(entry)
    return {"headings": headings, "paragraph_count": len(paras)}


def read_heading_section(docx_path: str | Path, heading_anchor_id: str | None = None, heading_text: str | None = None, max_chars: int = 20000) -> dict[str, Any]:
    outline = get_outline(docx_path)
    match = None
    for h in outline["headings"]:
        if heading_anchor_id and h["anchor_id"] == heading_anchor_id:
            match = h
            break
        if heading_text and h["text"] == heading_text:
            match = h
            break
    if match is None:
        raise ValueError("Heading not found")
    tree = load_document_tree(docx_path)
    paras = tree.getroot().xpath("//w:body//w:p", namespaces=NS)
    start = int(match["paragraph_index"])
    end = int(match["end_paragraph_index"])
    text_lines = [paragraph_text(p) for p in paras[start - 1 : end]]
    text = "\n".join(text_lines)
    truncated = len(text) > max_chars
    return {"heading": match, "text": text[:max_chars], "truncated": truncated, "char_count": len(text)}


def get_content_control_text(docx_path: str | Path, tag: str) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    sdt = find_sdt_by_tag(tree.getroot(), tag)
    if sdt is None:
        raise ValueError(f"Content control tag not found: {tag}")
    paragraphs = [paragraph_text(p) for p in sdt.xpath(".//w:p", namespaces=NS)]
    text = "\n".join(paragraphs) if paragraphs else "".join(t.text or "" for t in sdt.xpath(".//w:t", namespaces=NS))
    return {
        "tag": tag,
        "alias": sdt_alias(sdt),
        "id": sdt_id(sdt),
        "text": text,
        "text_sha256": _sha(text),
        "path": _element_path(sdt),
        "complexity": _complex_summary(sdt),
    }


def read_anchor(docx_path: str | Path, anchor_id: str, max_chars: int = 20000) -> dict[str, Any]:
    if anchor_id.startswith("cc:"):
        tag = anchor_id[3:]
        data = get_content_control_text(docx_path, tag)
        data["anchor_id"] = anchor_id
        data["truncated"] = len(data["text"]) > max_chars
        data["text"] = data["text"][:max_chars]
        return data
    if anchor_id.startswith("heading:"):
        return read_heading_section(docx_path, heading_anchor_id=anchor_id, max_chars=max_chars)
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    if anchor_id.startswith("p:"):
        pid = anchor_id[2:]
        found = root.xpath("//w:p[@w14:paraId=$pid]", pid=pid, namespaces=NS)
        if not found:
            raise ValueError(f"Paragraph anchor not found: {anchor_id}")
        text = paragraph_text(found[0])
        return {"anchor_id": anchor_id, "text": text[:max_chars], "text_sha256": _sha(text), "truncated": len(text) > max_chars, "path": _element_path(found[0])}
    if anchor_id.startswith("bookmark:"):
        name = anchor_id[len("bookmark:") :]
        bm = root.xpath("//w:bookmarkStart[@w:name=$name]", name=name, namespaces=NS)
        if not bm:
            raise ValueError(f"Bookmark not found: {name}")
        return {"anchor_id": anchor_id, "path": _element_path(bm[0]), "text": "", "note": "Bookmark location returned; range extraction is intentionally not used for safe editing."}
    raise ValueError(f"Unsupported anchor_id: {anchor_id}")


def find_sdt_by_tag(root: etree._Element, tag: str) -> etree._Element | None:
    candidates = root.xpath("//w:sdt[w:sdtPr/w:tag[@w:val=$tag]]", tag=tag, namespaces=NS)
    return candidates[0] if candidates else None


def _new_text_run(text: str, rpr: etree._Element | None = None) -> etree._Element:
    r = etree.Element(qn("w:r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, qn("w:t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def _new_paragraph(text: str, ppr: etree._Element | None = None, rpr: etree._Element | None = None) -> etree._Element:
    p = etree.Element(qn("w:p"))
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    p.append(_new_text_run(text, rpr))
    return p


def _replace_paragraph_text(p: etree._Element, new_text: str, preserve_run_style: bool = True, allow_complex_content: bool = False) -> None:
    if not allow_complex_content and _has_complex_content(p):
        raise ValueError(f"Refusing to replace paragraph containing complex Word objects at {_element_path(p)}")
    ppr = p.find("./w:pPr", namespaces=NS)
    first_rpr = p.find(".//w:rPr", namespaces=NS) if preserve_run_style else None
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    p.append(_new_text_run(new_text, first_rpr))


def _replace_paragraphs_text(
    paragraphs: list[etree._Element],
    text: str,
    *,
    allow_paragraph_count_change: bool = False,
    preserve_run_style: bool = True,
    allow_complex_content: bool = False,
) -> None:
    if not paragraphs:
        raise ValueError("No paragraphs available for text replacement")
    lines = text.splitlines() or [""]
    if not allow_paragraph_count_change and len(lines) != len(paragraphs):
        raise ValueError(
            f"Paragraph count would change from {len(paragraphs)} to {len(lines)}. "
            "Set allow_paragraph_count_change=true or use append/prepend tools explicitly."
        )
    if not allow_complex_content:
        for p in paragraphs:
            if _has_complex_content(p):
                raise ValueError(f"Refusing to replace paragraph containing complex Word objects at {_element_path(p)}")
    if len(lines) == len(paragraphs):
        for p, line in zip(paragraphs, lines):
            _replace_paragraph_text(p, line, preserve_run_style=preserve_run_style, allow_complex_content=allow_complex_content)
        return
    parent = paragraphs[0].getparent()
    first_idx = parent.index(paragraphs[0])
    ppr = paragraphs[0].find("./w:pPr", namespaces=NS)
    rpr = paragraphs[0].find(".//w:rPr", namespaces=NS) if preserve_run_style else None
    for p in paragraphs:
        parent.remove(p)
    for offset, line in enumerate(lines):
        parent.insert(first_idx + offset, _new_paragraph(line, ppr, rpr))


def _replace_sdt_text(
    sdt: etree._Element,
    text: str,
    *,
    allow_paragraph_count_change: bool = False,
    preserve_run_style: bool = True,
    allow_complex_content: bool = False,
) -> None:
    content = sdt.find("./w:sdtContent", namespaces=NS)
    if content is None:
        raise ValueError("Malformed content control: missing w:sdtContent")
    if not allow_complex_content and _complex_summary(sdt)["tables"] + _complex_summary(sdt)["drawings"] + _complex_summary(sdt)["fields"] + _complex_summary(sdt)["comments"] + _complex_summary(sdt)["tracked_changes"]:
        raise ValueError(f"Refusing to replace complex content control {sdt_tag(sdt)}. Use targeted table/comment tools or allow_complex_content=true.")
    paragraphs = _direct_paragraphs(content)
    if paragraphs:
        _replace_paragraphs_text(
            paragraphs,
            text,
            allow_paragraph_count_change=allow_paragraph_count_change,
            preserve_run_style=preserve_run_style,
            allow_complex_content=allow_complex_content,
        )
    else:
        p = _new_paragraph(text)
        content.append(p)


def _append_or_prepend_sdt_text(sdt: etree._Element, text: str, mode: str) -> None:
    content = sdt.find("./w:sdtContent", namespaces=NS)
    if content is None:
        raise ValueError("Malformed content control: missing w:sdtContent")
    if _complex_summary(sdt)["tables"] or _complex_summary(sdt)["drawings"]:
        raise ValueError("Appending/prepending text to a content control containing tables/images is blocked; target a simpler text content control.")
    paragraphs = _direct_paragraphs(content)
    lines = text.splitlines() or [""]
    template_p = paragraphs[-1] if paragraphs else None
    ppr = template_p.find("./w:pPr", namespaces=NS) if template_p is not None else None
    rpr = template_p.find(".//w:rPr", namespaces=NS) if template_p is not None else None
    new_paras = [_new_paragraph(line, ppr, rpr) for line in lines]
    if mode == "append":
        for p in new_paras:
            content.append(p)
    else:
        insert_at = 0
        for offset, p in enumerate(new_paras):
            content.insert(insert_at + offset, p)


def _replace_text_in_sdt(sdt: etree._Element, find: str, replace: str, occurrence: str = "all") -> int:
    if _complex_summary(sdt)["fields"] or _complex_summary(sdt)["tracked_changes"]:
        raise ValueError("Target content control contains fields or tracked changes; find/replace is blocked.")
    count = 0
    for t in sdt.xpath(".//w:t", namespaces=NS):
        if t.text and find in t.text:
            if occurrence == "first" and count:
                break
            n = t.text.count(find)
            t.text = t.text.replace(find, replace, 1 if occurrence == "first" else -1)
            count += 1 if occurrence == "first" else n
            if occurrence == "first":
                break
    return count


def _find_paragraph(root: etree._Element, *, para_id: str | None = None, paragraph_index: int | None = None) -> etree._Element | None:
    if para_id:
        found = root.xpath("//w:p[@w14:paraId=$pid]", pid=str(para_id), namespaces=NS)
        return found[0] if found else None
    if paragraph_index is not None:
        paras = root.xpath("//w:body//w:p", namespaces=NS)
        n = int(paragraph_index)
        if 1 <= n <= len(paras):
            return paras[n - 1]
    return None


def _tables_in_scope(root: etree._Element, scope_tag: str | None = None) -> list[etree._Element]:
    if scope_tag:
        sdt = find_sdt_by_tag(root, scope_tag)
        if sdt is None:
            raise ValueError(f"Content control tag not found for table scope: {scope_tag}")
        return sdt.xpath(".//w:tbl", namespaces=NS)
    return root.xpath("//w:body//w:tbl", namespaces=NS)


def _table_cell(tbl: etree._Element, row: int, col: int) -> etree._Element:
    rows = tbl.xpath("./w:tr", namespaces=NS)
    if row < 1 or row > len(rows):
        raise ValueError("row out of range")
    cells = rows[row - 1].xpath("./w:tc", namespaces=NS)
    if col < 1 or col > len(cells):
        raise ValueError("col out of range")
    return cells[col - 1]


def _cell_text(tc: etree._Element) -> str:
    return "\n".join(paragraph_text(p) for p in tc.xpath("./w:p", namespaces=NS))


def _table_text(tbl: etree._Element) -> str:
    rows: list[str] = []
    for tr in tbl.xpath("./w:tr", namespaces=NS):
        rows.append("\t".join(_cell_text(tc) for tc in tr.xpath("./w:tc", namespaces=NS)))
    return "\n".join(rows)


def list_tables(docx_path: str | Path, max_cell_chars: int = 120) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    tables = root.xpath("//w:body//w:tbl", namespaces=NS)
    out = []
    for ti, tbl in enumerate(tables, start=1):
        rows = tbl.xpath("./w:tr", namespaces=NS)
        row_summaries = []
        for ri, tr in enumerate(rows[:5], start=1):
            cells = tr.xpath("./w:tc", namespaces=NS)
            row_summaries.append([_cell_text(tc).replace("\n", " / ")[:max_cell_chars] for tc in cells])
        out.append(
            {
                "table_index": ti,
                "path": _element_path(tbl),
                "row_count": len(rows),
                "column_counts": [len(r.xpath("./w:tc", namespaces=NS)) for r in rows],
                "content_control_tag": _ancestor_sdt_tag(tbl),
                "preview_rows": row_summaries,
                "text_sha256": _sha(_table_text(tbl)),
                "xml_sha256": _canonical_hash(tbl),
                "complexity": _complex_summary(tbl),
            }
        )
    return {"tables": out, "table_count": len(out)}


def read_table(docx_path: str | Path, table_index: int, scope_tag: str | None = None, max_chars_per_cell: int = 2000) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    tables = _tables_in_scope(root, scope_tag)
    if table_index < 1 or table_index > len(tables):
        raise ValueError("table_index out of range")
    tbl = tables[table_index - 1]
    rows_out = []
    for ri, tr in enumerate(tbl.xpath("./w:tr", namespaces=NS), start=1):
        cells = []
        for ci, tc in enumerate(tr.xpath("./w:tc", namespaces=NS), start=1):
            text = _cell_text(tc)
            cells.append({"row": ri, "col": ci, "text": text[:max_chars_per_cell], "text_sha256": _sha(text), "truncated": len(text) > max_chars_per_cell})
        rows_out.append(cells)
    return {"table_index": table_index, "scope_tag": scope_tag, "rows": rows_out, "text_sha256": _sha(_table_text(tbl)), "path": _element_path(tbl)}


def find_text(docx_path: str | Path, query: str, case_sensitive: bool = False, max_results: int = 50, context_chars: int = 120) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    q = query if case_sensitive else query.lower()
    results = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        text = paragraph_text(p)
        hay = text if case_sensitive else text.lower()
        pos = hay.find(q)
        if pos >= 0:
            start = max(0, pos - context_chars)
            end = min(len(text), pos + len(query) + context_chars)
            results.append(
                {
                    "kind": "paragraph",
                    "paragraph_index": idx,
                    "paraId": p.get(qn("w14:paraId")),
                    "content_control_tag": _ancestor_sdt_tag(p),
                    "table_index": _ancestor_table_index(p, root),
                    "path": _element_path(p),
                    "context": text[start:end],
                    "text_sha256": _sha(text),
                }
            )
            if len(results) >= max_results:
                break
    return {"query": query, "count": len(results), "results": results}


def _extract_plain_text(docx_path: str | Path) -> str:
    tree = load_document_tree(docx_path)
    return "\n".join(paragraph_text(p) for p in tree.getroot().xpath("//w:body//w:p", namespaces=NS))


def diff_text(source_docx: str | Path, target_docx: str | Path, context: int = 2) -> str:
    a = _extract_plain_text(source_docx).splitlines()
    b = _extract_plain_text(target_docx).splitlines()
    return "\n".join(difflib.unified_diff(a, b, fromfile=str(source_docx), tofile=str(target_docx), n=context))


def _text_hashes_by_content_control(root: etree._Element) -> dict[str, str]:
    out = {}
    for sdt in root.xpath("//w:sdt", namespaces=NS):
        tag = sdt_tag(sdt)
        if tag:
            out[tag] = _canonical_hash(sdt)
    return out


def _paragraph_hashes_by_para_id(root: etree._Element) -> dict[str, dict[str, Any]]:
    out = {}
    for p in root.xpath("//w:body//w:p", namespaces=NS):
        pid = p.get(qn("w14:paraId"))
        if pid:
            out[pid] = {"hash": _canonical_hash(p), "content_control_tag": _ancestor_sdt_tag(p), "table_index": _ancestor_table_index(p, root)}
    return out


def _table_hashes(root: etree._Element) -> dict[int, dict[str, Any]]:
    out = {}
    tables = root.xpath("//w:body//w:tbl", namespaces=NS)
    for idx, tbl in enumerate(tables, start=1):
        rows = tbl.xpath("./w:tr", namespaces=NS)
        out[idx] = {
            "hash": _canonical_hash(tbl),
            "row_count": len(rows),
            "column_counts": [len(r.xpath("./w:tc", namespaces=NS)) for r in rows],
            "content_control_tag": _ancestor_sdt_tag(tbl),
        }
    return out


def _paragraph_global_index(root: etree._Element, paragraph: etree._Element) -> int | None:
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        if p is paragraph:
            return idx
    return None


def _table_cell_hashes(root: etree._Element) -> dict[int, dict[str, Any]]:
    out: dict[int, dict[str, Any]] = {}
    for table_index, tbl in enumerate(root.xpath("//w:body//w:tbl", namespaces=NS), start=1):
        cells: dict[str, str] = {}
        column_counts: list[int] = []
        rows = tbl.xpath("./w:tr", namespaces=NS)
        for r_idx, row in enumerate(rows, start=1):
            row_cells = row.xpath("./w:tc", namespaces=NS)
            column_counts.append(len(row_cells))
            for c_idx, cell in enumerate(row_cells, start=1):
                cells[f"{r_idx}:{c_idx}"] = _canonical_hash(cell)
        out[table_index] = {"row_count": len(rows), "column_counts": column_counts, "cells": cells}
    return out


def _normalise_table_cell_refs(refs: Iterable[Any] | None) -> set[tuple[int, int, int]]:
    out: set[tuple[int, int, int]] = set()
    for ref in refs or []:
        if isinstance(ref, dict):
            out.add((int(ref["table_index"]), int(ref["row"]), int(ref["col"])))
        elif isinstance(ref, (list, tuple)) and len(ref) == 3:
            out.add((int(ref[0]), int(ref[1]), int(ref[2])))
        elif isinstance(ref, str):
            parts = ref.replace(",", ":").split(":")
            if len(parts) != 3:
                raise ValueError(f"Invalid table cell reference: {ref!r}; expected table:row:col")
            out.add((int(parts[0]), int(parts[1]), int(parts[2])))
        else:
            raise ValueError(f"Invalid table cell reference: {ref!r}")
    return out


def _body_block_sequence(root: etree._Element) -> list[dict[str, Any]]:
    body = root.find(".//w:body", namespaces=NS)
    if body is None:
        return []
    table_paths = {tbl.getroottree().getpath(tbl): idx for idx, tbl in enumerate(root.xpath("//w:body//w:tbl", namespaces=NS), start=1)}
    paragraph_paths = {par.getroottree().getpath(par): idx for idx, par in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1)}
    seq: list[dict[str, Any]] = []
    for block_index, child in enumerate(list(body), start=1):
        local = etree.QName(child).localname
        if local == "sectPr":
            continue
        descendant_sdt_tags = sorted({tag for tag in (sdt_tag(sdt) for sdt in child.xpath(".//w:sdt", namespaces=NS)) if tag})
        if local == "sdt":
            own = sdt_tag(child)
            if own and own not in descendant_sdt_tags:
                descendant_sdt_tags.insert(0, own)
        table_nodes = list(child.xpath(".//w:tbl", namespaces=NS))
        if local == "tbl":
            table_nodes.insert(0, child)
        table_indices = sorted({table_paths[tbl.getroottree().getpath(tbl)] for tbl in table_nodes if tbl.getroottree().getpath(tbl) in table_paths})
        paragraph_nodes = list(child.xpath(".//w:p", namespaces=NS))
        if local == "p":
            paragraph_nodes.insert(0, child)
        paragraph_indices = sorted({paragraph_paths[par.getroottree().getpath(par)] for par in paragraph_nodes if par.getroottree().getpath(par) in paragraph_paths})
        seq.append({
            "block_index": block_index,
            "kind": local,
            "xml_sha256": _canonical_hash(child),
            "text_sha256": _sha(paragraph_text(child) if local == "p" else "\n".join(paragraph_text(p) for p in child.xpath(".//w:p", namespaces=NS))),
            "content_control_tags": descendant_sdt_tags,
            "table_indices": table_indices,
            "paragraph_indices": paragraph_indices,
        })
    return seq


def _sequence_preservation_report(source_blocks: list[dict[str, Any]], target_blocks: list[dict[str, Any]], *,
                                  touched_content_control_tags: set[str],
                                  touched_table_indices: set[int],
                                  touched_paragraph_indices: set[int]) -> dict[str, Any]:
    def is_touched(block: dict[str, Any]) -> bool:
        return bool(
            set(block.get("content_control_tags") or []) & touched_content_control_tags
            or set(int(i) for i in (block.get("table_indices") or [])) & touched_table_indices
            or set(int(i) for i in (block.get("paragraph_indices") or [])) & touched_paragraph_indices
        )

    source_protected = [b for b in source_blocks if not is_touched(b)]
    target_protected = [b for b in target_blocks if not is_touched(b)]
    j = 0
    missing: list[dict[str, Any]] = []
    for sb in source_protected:
        while j < len(target_protected) and target_protected[j]["xml_sha256"] != sb["xml_sha256"]:
            j += 1
        if j >= len(target_protected):
            missing.append({
                "block_index": sb.get("block_index"),
                "kind": sb.get("kind"),
                "text_sha256": sb.get("text_sha256"),
                "content_control_tags": sb.get("content_control_tags"),
                "table_indices": sb.get("table_indices"),
                "paragraph_indices": sb.get("paragraph_indices"),
            })
        else:
            j += 1
    return {
        "ok": not missing,
        "source_block_count": len(source_blocks),
        "target_block_count": len(target_blocks),
        "protected_source_block_count": len(source_protected),
        "protected_target_block_count": len(target_protected),
        "missing_or_modified_protected_blocks": missing[:50],
    }


def _structural_fingerprint(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    return {
        "content_controls": _text_hashes_by_content_control(root),
        "paragraphs": _paragraph_hashes_by_para_id(root),
        "tables": _table_hashes(root),
        "table_cells": _table_cell_hashes(root),
        "body_blocks": _body_block_sequence(root),
    }


def _apply_add_comment(root: etree._Element, target_p: etree._Element, comment_text: str, author: str, initials: str, existing_comments: bytes | None) -> tuple[bytes, int]:
    if existing_comments:
        comments_tree = _parse_xml(existing_comments)
        comments_root = comments_tree.getroot()
    else:
        comments_root = etree.Element(qn("w:comments"), nsmap={"w": W_NS})
        comments_tree = comments_root.getroottree()
    existing_ids = [int(c.get(qn("w:id"))) for c in comments_root.xpath("//w:comment", namespaces=NS) if c.get(qn("w:id"), "").isdigit()]
    cid = max(existing_ids, default=-1) + 1

    # Place comment markers around the visible paragraph content. This intentionally changes only the target paragraph.
    start = etree.Element(qn("w:commentRangeStart"))
    start.set(qn("w:id"), str(cid))
    end = etree.Element(qn("w:commentRangeEnd"))
    end.set(qn("w:id"), str(cid))
    insert_start = 1 if target_p.find("./w:pPr", namespaces=NS) is not None else 0
    target_p.insert(insert_start, start)
    ref_run = etree.Element(qn("w:r"))
    rpr = etree.SubElement(ref_run, qn("w:rPr"))
    rstyle = etree.SubElement(rpr, qn("w:rStyle"))
    rstyle.set(qn("w:val"), "CommentReference")
    cref = etree.SubElement(ref_run, qn("w:commentReference"))
    cref.set(qn("w:id"), str(cid))
    target_p.append(end)
    target_p.append(ref_run)

    comment = etree.SubElement(comments_root, qn("w:comment"))
    comment.set(qn("w:id"), str(cid))
    comment.set(qn("w:author"), author)
    comment.set(qn("w:initials"), initials)
    comment.set(qn("w:date"), datetime.now(timezone.utc).replace(microsecond=0).isoformat())
    p = etree.SubElement(comment, qn("w:p"))
    r = etree.SubElement(p, qn("w:r"))
    t = etree.SubElement(r, qn("w:t"))
    t.text = comment_text
    return _serialize_xml(comments_tree), cid


def _ensure_comments_relationship_and_content_type(docx_path: Path, replacements: dict[str, bytes]) -> None:
    # document.xml.rels
    if _zip_has(docx_path, DOCUMENT_RELS):
        rels_tree = _parse_xml(_zip_read(docx_path, DOCUMENT_RELS))
        rels_root = rels_tree.getroot()
    else:
        rels_root = etree.Element(f"{{{PKG_REL_NS}}}Relationships", nsmap={None: PKG_REL_NS})
        rels_tree = rels_root.getroottree()
    has_rel = any(rel.get("Type") == COMMENTS_REL_TYPE for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"))
    if not has_rel:
        ids = [rel.get("Id") or "" for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")]
        n = 1
        while f"rId{n}" in ids:
            n += 1
        rel = etree.SubElement(rels_root, f"{{{PKG_REL_NS}}}Relationship")
        rel.set("Id", f"rId{n}")
        rel.set("Type", COMMENTS_REL_TYPE)
        rel.set("Target", "comments.xml")
    replacements[DOCUMENT_RELS] = _serialize_xml(rels_tree)

    # Content types
    ct_tree = _parse_xml(_zip_read(docx_path, CONTENT_TYPES))
    ct_root = ct_tree.getroot()
    has_override = any(el.get("PartName") == "/word/comments.xml" for el in ct_root.findall(f"{{{CT_NS}}}Override"))
    if not has_override:
        override = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", COMMENTS_CONTENT_TYPE)
    replacements[CONTENT_TYPES] = _serialize_xml(ct_tree)


def _target_paragraph_from_op(root: etree._Element, op: dict[str, Any]) -> etree._Element:
    # For most paragraph-scoped operations, `tag` is a content-control locator.
    # For wrap_paragraph_with_content_control, however, `tag` is the *new* tag,
    # so it must not be resolved as an existing content control.
    if op.get("op") == "wrap_paragraph_with_content_control":
        tag = op.get("content_control_tag") or op.get("target_tag")
    else:
        tag = op.get("tag") or op.get("content_control_tag")
    if tag:
        sdt = find_sdt_by_tag(root, str(tag))
        if sdt is None:
            raise ValueError(f"Content control tag not found: {tag}")
        paragraphs = sdt.xpath(".//w:p", namespaces=NS)
        if not paragraphs:
            raise ValueError(f"Content control has no paragraph: {tag}")
        return paragraphs[0]
    p = _find_paragraph(root, para_id=op.get("paraId"), paragraph_index=op.get("paragraph_index"))
    if p is None:
        raise ValueError("Target paragraph not found")
    return p


def _wrap_paragraph_with_sdt(p: etree._Element, tag: str, alias: str | None = None, lock: bool = True) -> etree._Element:
    parent = p.getparent()
    if parent is None:
        raise ValueError("Paragraph has no parent")
    if _ancestor_sdt_tag(p):
        raise ValueError("Paragraph is already inside a content control")
    idx = parent.index(p)
    parent.remove(p)
    sdt = etree.Element(qn("w:sdt"))
    sdt_pr = etree.SubElement(sdt, qn("w:sdtPr"))
    alias_el = etree.SubElement(sdt_pr, qn("w:alias"))
    alias_el.set(qn("w:val"), alias or tag)
    tag_el = etree.SubElement(sdt_pr, qn("w:tag"))
    tag_el.set(qn("w:val"), tag)
    id_el = etree.SubElement(sdt_pr, qn("w:id"))
    id_el.set(qn("w:val"), str(int(hashlib.sha1(tag.encode("utf-8")).hexdigest()[:7], 16)))
    if lock:
        lock_el = etree.SubElement(sdt_pr, qn("w:lock"))
        lock_el.set(qn("w:val"), "sdtLocked")
    content = etree.SubElement(sdt, qn("w:sdtContent"))
    content.append(p)
    parent.insert(idx, sdt)
    return sdt


def assess_patchset(docx_path: str | Path, patchset: dict[str, Any]) -> dict[str, Any]:
    docx_path = Path(docx_path)
    operations = patchset.get("operations") or []
    guard = patchset.get("guard") or {}
    require_preconditions = bool(guard.get("require_preconditions", False))
    risks: list[dict[str, Any]] = []
    if not operations:
        raise ValueError("patchset.operations must contain at least one operation")
    if require_preconditions and not patchset.get("source_sha256"):
        risks.append({"severity": "error", "code": "missing_source_sha256", "message": "guard.require_preconditions=true requires patchset.source_sha256."})
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    touched_tags: set[str] = set()
    touched_para_ids: set[str] = set()
    touched_paragraph_indices: set[int] = set()
    touched_table_indices: set[int] = set()
    touched_table_cells: set[tuple[int, int, int]] = set()
    requires_structural_change = False

    for idx, op in enumerate(operations):
        typ = op.get("op")
        try:
            if typ in {"replace_content_control_text", "append_content_control_text", "prepend_content_control_text", "replace_text_in_content_control"}:
                tag = op.get("tag")
                if not tag:
                    raise ValueError("tag is required")
                sdt = find_sdt_by_tag(root, str(tag))
                if sdt is None:
                    raise ValueError(f"content control tag not found: {tag}")
                touched_tags.add(str(tag))
                complexity = _complex_summary(sdt)
                if any(complexity.values()) and not op.get("allow_complex_content", False) and typ == "replace_content_control_text":
                    risks.append({"severity": "error", "operation_index": idx, "code": "complex_content_control", "message": "Replacement target contains complex Word objects", "complexity": complexity})
                if typ in {"append_content_control_text", "prepend_content_control_text"}:
                    requires_structural_change = True
                if op.get("expected_old_text") is None and op.get("expected_old_sha256") is None:
                    risks.append({"severity": "error" if require_preconditions else "warning", "operation_index": idx, "code": "missing_precondition", "message": "No expected_old_text/expected_old_sha256 supplied; concurrent edits may be missed."})
            elif typ in {"replace_paragraph_text", "insert_paragraph_after", "insert_paragraph_before", "add_comment", "wrap_paragraph_with_content_control"}:
                p = _target_paragraph_from_op(root, op)
                pid = p.get(qn("w14:paraId"))
                pidx = _paragraph_global_index(root, p)
                if pid:
                    touched_para_ids.add(pid)
                if pidx:
                    touched_paragraph_indices.add(pidx)
                if typ.startswith("insert_") or typ in {"add_comment", "wrap_paragraph_with_content_control"}:
                    requires_structural_change = True
                if typ == "wrap_paragraph_with_content_control":
                    tag = op.get("tag")
                    if not tag:
                        raise ValueError("tag is required")
                    if find_sdt_by_tag(root, str(tag)) is not None:
                        raise ValueError(f"content control tag already exists: {tag}")
                    touched_tags.add(str(tag))
                if _has_complex_content(p) and typ == "replace_paragraph_text" and not op.get("allow_complex_content", False):
                    risks.append({"severity": "error", "operation_index": idx, "code": "complex_paragraph", "message": "Target paragraph contains complex Word objects."})
                if op.get("expected_old_text") is None and op.get("expected_old_sha256") is None:
                    risks.append({"severity": "error" if require_preconditions or typ in HIGH_RISK_OPS else "warning", "operation_index": idx, "code": "missing_precondition", "message": "No expected_old_text/expected_old_sha256 supplied for paragraph-scoped operation."})
            elif typ in {"replace_table_cell_text", "append_table_row"}:
                scope_tag = op.get("scope_tag")
                tables = _tables_in_scope(root, scope_tag)
                ti = int(op.get("table_index", 1))
                if ti < 1 or ti > len(tables):
                    raise ValueError("table_index out of range")
                tbl = tables[ti - 1]
                global_index = _ancestor_table_index(tbl, root) or ti
                touched_table_indices.add(global_index)
                if typ == "append_table_row":
                    requires_structural_change = True
                if typ == "replace_table_cell_text":
                    row = int(op.get("row", 1))
                    col = int(op.get("col", 1))
                    tc = _table_cell(tbl, row, col)
                    touched_table_cells.add((global_index, row, col))
                    if _has_complex_content(tc) and not op.get("allow_complex_content", False):
                        risks.append({"severity": "error", "operation_index": idx, "code": "complex_table_cell", "message": "Target table cell contains complex Word objects."})
                if op.get("expected_old_text") is None and op.get("expected_old_sha256") is None:
                    target = "table" if typ == "append_table_row" else "table cell"
                    risks.append({"severity": "error" if require_preconditions or typ in HIGH_RISK_OPS else "warning", "operation_index": idx, "code": "missing_precondition", "message": f"No expected_old_text/expected_old_sha256 supplied for {target} operation."})
            else:
                risks.append({"severity": "error", "operation_index": idx, "code": "unsupported_op", "message": f"Unsupported op: {typ}"})
        except Exception as exc:
            risks.append({"severity": "error", "operation_index": idx, "code": "target_resolution_failed", "message": str(exc)})
    return {
        "ok": not any(r["severity"] == "error" for r in risks),
        "operation_count": len(operations),
        "requires_structural_change": requires_structural_change,
        "touched": {
            "content_control_tags": sorted(touched_tags),
            "paraIds": sorted(touched_para_ids),
            "paragraph_indices": sorted(touched_paragraph_indices),
            "table_indices": sorted(touched_table_indices),
            "table_cells": sorted(f"{t}:{r}:{c}" for t, r, c in touched_table_cells),
        },
        "risks": risks,
    }


def apply_patchset(docx_path: str | Path, patchset: dict[str, Any], output_path: str | Path | None = None) -> dict[str, Any]:
    """Apply constrained Word edits. Default policy is structure-protection first.

    All ordinary writes modify only word/document.xml; comment operations may add/update
    comments.xml, document relationships, and content types. The source DOCX is never
    overwritten unless the caller explicitly chooses the same output path, which is not
    recommended.
    """
    docx_path = Path(docx_path)
    source_sha = patchset.get("source_sha256")
    if source_sha and source_sha != _sha(docx_path.read_bytes()):
        raise ValueError("patchset.source_sha256 does not match current DOCX; refusing stale write")
    if output_path is None:
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        output_path = docx_path.with_name(f"{docx_path.stem}.wordai.{stamp}.docx")
    output_path = Path(output_path)
    guard = patchset.get("guard") or {}
    allow_overwrite = bool(guard.get("allow_overwrite", False) or patchset.get("overwrite", False))
    if output_path.resolve() == docx_path.resolve() and not allow_overwrite:
        raise ValueError("Refusing to overwrite the source DOCX. Use a distinct output_path or set guard.allow_overwrite=true after creating a backup.")
    if output_path.exists() and not allow_overwrite:
        raise FileExistsError(f"output_path already exists: {output_path}; set patchset.overwrite=true or guard.allow_overwrite=true after creating a backup")

    strict = patchset.get("strict", True)
    operations = patchset.get("operations") or []
    if patchset.get("schema_version") not in {"1.0", "1.1", "2.0", None}:
        raise ValueError("Unsupported patchset.schema_version")
    if not operations:
        raise ValueError("patchset.operations must contain at least one operation")

    assessment = assess_patchset(docx_path, patchset)
    if not assessment["ok"]:
        raise ValueError("PatchSet failed safety assessment: " + json.dumps(assessment["risks"], ensure_ascii=False))

    before = inspect_docx(docx_path)
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    styles = _load_paragraph_styles(docx_path)
    applied: list[dict[str, Any]] = []
    replacements: dict[str, bytes] = {}
    touched_tags: set[str] = set()
    touched_para_ids: set[str] = set()
    touched_paragraph_indices: set[int] = set()
    touched_table_indices: set[int] = set()
    touched_table_cells: set[tuple[int, int, int]] = set()
    allowed_count_changes: set[str] = set()
    allowed_part_changes: set[str] = {DOCUMENT_XML}
    allow_table_dimension_change = False
    allow_paragraph_count_change_for_validation = False
    allowed_added_content_control_tags: set[str] = set()

    for idx, op in enumerate(operations):
        typ = op.get("op")
        if typ == "replace_content_control_text":
            tag = op.get("tag")
            text = op.get("text")
            if not isinstance(tag, str) or text is None:
                raise ValueError(f"operation {idx}: tag and text are required")
            sdt = find_sdt_by_tag(root, tag)
            if sdt is None:
                raise ValueError(f"operation {idx}: content control tag not found: {tag}")
            current = get_content_control_text_from_tree(sdt)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} content control {tag}")
            allow_pc = bool(op.get("allow_paragraph_count_change", False))
            _replace_sdt_text(
                sdt,
                str(text),
                allow_paragraph_count_change=allow_pc,
                preserve_run_style=bool(op.get("preserve_style", True)),
                allow_complex_content=bool(op.get("allow_complex_content", False)),
            )
            if allow_pc:
                allowed_count_changes.add("paragraph_count")
                allow_paragraph_count_change_for_validation = True
            touched_tags.add(tag)
            applied.append({"op": typ, "tag": tag, "path": _element_path(sdt), "old_sha256": _sha(current), "new_sha256": _sha(str(text))})
        elif typ in {"append_content_control_text", "prepend_content_control_text"}:
            tag = op.get("tag")
            text = op.get("text")
            if not isinstance(tag, str) or text is None:
                raise ValueError(f"operation {idx}: tag and text are required")
            sdt = find_sdt_by_tag(root, tag)
            if sdt is None:
                raise ValueError(f"operation {idx}: content control tag not found: {tag}")
            current = get_content_control_text_from_tree(sdt)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} content control {tag}")
            _append_or_prepend_sdt_text(sdt, str(text), "append" if typ.startswith("append") else "prepend")
            touched_tags.add(tag)
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change_for_validation = True
            applied.append({"op": typ, "tag": tag, "path": _element_path(sdt), "old_sha256": _sha(current), "inserted_sha256": _sha(str(text))})
        elif typ == "replace_text_in_content_control":
            tag = op.get("tag")
            find = op.get("find")
            replace = op.get("replace")
            if not isinstance(tag, str) or find is None or replace is None:
                raise ValueError(f"operation {idx}: tag, find and replace are required")
            sdt = find_sdt_by_tag(root, tag)
            if sdt is None:
                raise ValueError(f"operation {idx}: content control tag not found: {tag}")
            current = get_content_control_text_from_tree(sdt)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} content control {tag}")
            count = _replace_text_in_sdt(sdt, str(find), str(replace), str(op.get("occurrence", "all")))
            if count == 0 and bool(op.get("require_match", True)):
                raise ValueError(f"operation {idx}: find text not found")
            touched_tags.add(tag)
            applied.append({"op": typ, "tag": tag, "replace_count": count, "path": _element_path(sdt)})
        elif typ == "replace_paragraph_text":
            text = op.get("text")
            if text is None:
                raise ValueError(f"operation {idx}: text is required")
            p = _find_paragraph(root, para_id=op.get("paraId"), paragraph_index=op.get("paragraph_index"))
            if p is None:
                raise ValueError(f"operation {idx}: paragraph not found")
            pidx = _paragraph_global_index(root, p)
            current = paragraph_text(p)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} paragraph")
            _replace_paragraph_text(p, str(text), preserve_run_style=bool(op.get("preserve_style", True)), allow_complex_content=bool(op.get("allow_complex_content", False)))
            pid = p.get(qn("w14:paraId"))
            if pid:
                touched_para_ids.add(pid)
            if pidx:
                touched_paragraph_indices.add(pidx)
            applied.append({"op": typ, "path": _element_path(p), "paraId": pid, "paragraph_index": pidx, "old_sha256": _sha(current), "new_sha256": _sha(str(text))})
        elif typ in {"insert_paragraph_after", "insert_paragraph_before"}:
            text = op.get("text")
            if text is None:
                raise ValueError(f"operation {idx}: text is required")
            p = _target_paragraph_from_op(root, op)
            current = paragraph_text(p)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} anchor paragraph")
            parent = p.getparent()
            inherit_style = bool(op.get("inherit_style", True))
            # Structure-first default: never inherit heading styles for inserted body paragraphs
            # unless the caller explicitly allows it. This prevents accidental heading-count
            # changes when Codex inserts content before/after a heading anchor.
            if _heading_level(_style_id_for_paragraph(p), styles, paragraph=p) is not None and not bool(op.get("inherit_heading_style", False)):
                inherit_style = False
            ppr = p.find("./w:pPr", namespaces=NS) if inherit_style else None
            rpr = p.find(".//w:rPr", namespaces=NS) if inherit_style else None
            lines = str(text).splitlines() or [""]
            insert_idx = parent.index(p) + (1 if typ.endswith("after") else 0)
            for offset, line in enumerate(lines):
                parent.insert(insert_idx + offset, _new_paragraph(line, ppr, rpr))
            pid = p.get(qn("w14:paraId"))
            pidx = _paragraph_global_index(root, p)
            if pid:
                touched_para_ids.add(pid)
            if pidx:
                touched_paragraph_indices.add(pidx)
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change_for_validation = True
            applied.append({"op": typ, "anchor_paraId": pid, "anchor_paragraph_index": pidx, "old_sha256": _sha(current), "inserted_paragraph_count": len(lines), "inserted_sha256": _sha(str(text))})
        elif typ == "replace_table_cell_text":
            tbl_index = int(op.get("table_index", 1))
            row = int(op.get("row", 1))
            col = int(op.get("col", 1))
            text = str(op.get("text", ""))
            tables = _tables_in_scope(root, op.get("scope_tag"))
            if tbl_index < 1 or tbl_index > len(tables):
                raise ValueError(f"operation {idx}: table_index out of range")
            tbl = tables[tbl_index - 1]
            global_ti = _ancestor_table_index(tbl, root) or tbl_index
            cell = _table_cell(tbl, row, col)
            current = _cell_text(cell)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} table cell")
            paras = cell.xpath("./w:p", namespaces=NS)
            if not paras:
                cell.append(_new_paragraph(text))
                allowed_count_changes.add("paragraph_count")
                allow_paragraph_count_change_for_validation = True
            else:
                _replace_paragraphs_text(
                    paras,
                    text,
                    allow_paragraph_count_change=bool(op.get("allow_paragraph_count_change", False)),
                    preserve_run_style=bool(op.get("preserve_style", True)),
                    allow_complex_content=bool(op.get("allow_complex_content", False)),
                )
                if bool(op.get("allow_paragraph_count_change", False)):
                    allowed_count_changes.add("paragraph_count")
                    allow_paragraph_count_change_for_validation = True
            touched_table_indices.add(global_ti)
            touched_table_cells.add((global_ti, row, col))
            applied.append({"op": typ, "table_index": global_ti, "row": row, "col": col, "old_sha256": _sha(current), "new_sha256": _sha(text)})
        elif typ == "append_table_row":
            tbl_index = int(op.get("table_index", 1))
            values = op.get("values")
            if not isinstance(values, list):
                raise ValueError(f"operation {idx}: values array is required")
            tables = _tables_in_scope(root, op.get("scope_tag"))
            if tbl_index < 1 or tbl_index > len(tables):
                raise ValueError(f"operation {idx}: table_index out of range")
            tbl = tables[tbl_index - 1]
            global_ti = _ancestor_table_index(tbl, root) or tbl_index
            current_table_text = _table_text(tbl)
            _assert_expected_text(current_table_text, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} table")
            rows = tbl.xpath("./w:tr", namespaces=NS)
            if not rows:
                raise ValueError("Cannot append row to empty table")
            template_row_index = int(op.get("template_row", len(rows)))
            if template_row_index < 1 or template_row_index > len(rows):
                raise ValueError("template_row out of range")
            new_row = copy.deepcopy(rows[template_row_index - 1])
            cells = new_row.xpath("./w:tc", namespaces=NS)
            if len(values) > len(cells):
                raise ValueError("values has more items than table columns")
            for ci, val in enumerate(values):
                paras = cells[ci].xpath("./w:p", namespaces=NS)
                if not paras:
                    cells[ci].append(_new_paragraph(str(val)))
                else:
                    _replace_paragraphs_text(paras, str(val), allow_paragraph_count_change=True)
            tbl.append(new_row)
            touched_table_indices.add(global_ti)
            allow_table_dimension_change = True
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change_for_validation = True
            applied.append({"op": typ, "table_index": global_ti, "old_sha256": _sha(current_table_text), "appended_cells": len(values)})
        elif typ == "wrap_paragraph_with_content_control":
            tag = op.get("tag")
            if not isinstance(tag, str) or not tag:
                raise ValueError(f"operation {idx}: tag is required")
            if find_sdt_by_tag(root, tag) is not None:
                raise ValueError(f"operation {idx}: content control tag already exists: {tag}")
            p = _target_paragraph_from_op(root, op)
            current = paragraph_text(p)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} paragraph")
            pid = p.get(qn("w14:paraId"))
            pidx = _paragraph_global_index(root, p)
            sdt = _wrap_paragraph_with_sdt(p, tag, op.get("alias"), bool(op.get("lock", True)))
            if pid:
                touched_para_ids.add(pid)
            if pidx:
                touched_paragraph_indices.add(pidx)
            touched_tags.add(tag)
            allowed_added_content_control_tags.add(tag)
            allowed_count_changes.add("content_control_count")
            applied.append({"op": typ, "tag": tag, "alias": op.get("alias"), "target_paraId": pid, "target_paragraph_index": pidx, "old_sha256": _sha(current), "path": _element_path(sdt)})
        elif typ == "add_comment":
            text = op.get("text")
            if not isinstance(text, str) or not text:
                raise ValueError(f"operation {idx}: non-empty text is required")
            p = _target_paragraph_from_op(root, op)
            current = paragraph_text(p)
            _assert_expected_text(current, op.get("expected_old_text"), op.get("expected_old_sha256"), f"operation {idx} paragraph")
            pid = p.get(qn("w14:paraId"))
            pidx = _paragraph_global_index(root, p)
            current_comments = _zip_read(docx_path, COMMENTS_XML) if _zip_has(docx_path, COMMENTS_XML) else None
            comments_xml, cid = _apply_add_comment(root, p, text, str(op.get("author", "Word AI")), str(op.get("initials", "AI")), current_comments)
            replacements[COMMENTS_XML] = comments_xml
            _ensure_comments_relationship_and_content_type(docx_path, replacements)
            if pid:
                touched_para_ids.add(pid)
            if pidx:
                touched_paragraph_indices.add(pidx)
            allowed_part_changes.update({COMMENTS_XML, DOCUMENT_RELS, CONTENT_TYPES})
            allowed_count_changes.add("comment_count")
            allowed_count_changes.add("comment_reference_count")
            applied.append({"op": typ, "comment_id": cid, "target_paraId": pid, "target_paragraph_index": pidx, "old_sha256": _sha(current), "path": _element_path(p)})
        else:
            raise ValueError(f"operation {idx}: unsupported op {typ!r}")

    replacements[DOCUMENT_XML] = _serialize_xml(tree)
    candidate = output_path.with_suffix(output_path.suffix + f".candidate.{os.getpid()}")
    candidate.unlink(missing_ok=True)
    rewrite_zip_member(docx_path, candidate, replacements)
    report = validate_structure(
        docx_path,
        candidate,
        strict=strict,
        allowed_part_changes=allowed_part_changes,
        allowed_count_changes=allowed_count_changes,
        touched_content_control_tags=touched_tags,
        touched_para_ids=touched_para_ids,
        touched_paragraph_indices=touched_paragraph_indices,
        touched_table_indices=touched_table_indices,
        touched_table_cells=touched_table_cells,
        allow_table_dimension_change=allow_table_dimension_change,
        allow_paragraph_count_change=allow_paragraph_count_change_for_validation,
        allowed_added_content_control_tags=allowed_added_content_control_tags,
    )
    after = inspect_docx(candidate)
    audit = {
        "source_path": str(docx_path),
        "output_path": str(output_path),
        "candidate_path": str(candidate),
        "patchset_reason": patchset.get("reason"),
        "safety_assessment": assessment,
        "applied": applied,
        "validation": report.to_dict(),
        "before_metrics": {k: before[k] for k in before if k.endswith("_count")},
        "after_metrics": {k: after[k] for k in before if k.endswith("_count")},
        "changed_text_diff": diff_text(docx_path, candidate, context=2),
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    abort_on_validation_error = bool(patchset.get("abort_on_validation_error", True))
    keep_invalid_output = bool(patchset.get("keep_invalid_output", False))
    if abort_on_validation_error and not report.ok:
        audit_path = output_path.with_suffix(".invalid.audit.json")
        if keep_invalid_output:
            invalid_path = output_path.with_suffix(".invalid.docx")
            invalid_path.unlink(missing_ok=True)
            os.replace(candidate, invalid_path)
            audit["invalid_output_path"] = str(invalid_path)
        else:
            candidate.unlink(missing_ok=True)
        audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
        raise ValueError("Validation failed; final DOCX was not committed. See audit: " + str(audit_path))
    os.replace(candidate, output_path)
    audit["candidate_path"] = None
    if isinstance(audit.get("validation"), dict):
        audit["validation"]["target_path"] = str(output_path)
    audit["changed_text_diff"] = diff_text(docx_path, output_path, context=2)
    audit_path = output_path.with_suffix(".audit.json")
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    return audit


def get_content_control_text_from_tree(sdt: etree._Element) -> str:
    paragraphs = [paragraph_text(p) for p in sdt.xpath(".//w:p", namespaces=NS)]
    return "\n".join(paragraphs) if paragraphs else "".join(t.text or "" for t in sdt.xpath(".//w:t", namespaces=NS))


def dry_run_patchset(docx_path: str | Path, patchset: dict[str, Any], keep_output: bool = False) -> dict[str, Any]:
    docx_path = Path(docx_path)
    dry_dir = docx_path.parent / ".wordai" / "dryruns"
    dry_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    out = dry_dir / f"{docx_path.stem}.dryrun.{stamp}.docx"
    audit = apply_patchset(docx_path, patchset, out)
    audit["dry_run"] = True
    audit["kept_output"] = keep_output
    if not keep_output:
        try:
            out.unlink(missing_ok=True)
            out.with_suffix(".audit.json").unlink(missing_ok=True)
        except Exception:
            pass
        audit["output_path"] = None
    return audit


def validate_structure(
    source_docx: str | Path,
    target_docx: str | Path,
    strict: bool = True,
    *,
    allowed_part_changes: Iterable[str] | None = None,
    allowed_count_changes: Iterable[str] | None = None,
    touched_content_control_tags: Iterable[str] | None = None,
    touched_para_ids: Iterable[str] | None = None,
    touched_paragraph_indices: Iterable[int] | None = None,
    touched_table_indices: Iterable[int] | None = None,
    touched_table_cells: Iterable[Any] | None = None,
    allow_table_dimension_change: bool = False,
    allow_paragraph_count_change: bool = False,
    allowed_added_content_control_tags: Iterable[str] | None = None,
) -> ValidationReport:
    source_docx = Path(source_docx)
    target_docx = Path(target_docx)
    issues: list[ValidationIssue] = []
    metrics: dict[str, Any] = {}
    allowed_part_changes = set(allowed_part_changes or {DOCUMENT_XML})
    allowed_count_changes = set(allowed_count_changes or set())
    touched_content_control_tags = set(touched_content_control_tags or set())
    touched_para_ids = set(touched_para_ids or set())
    touched_paragraph_indices = {int(x) for x in (touched_paragraph_indices or set())}
    touched_table_indices = {int(x) for x in (touched_table_indices or set())}
    touched_table_cells = _normalise_table_cell_refs(touched_table_cells)
    allowed_added_content_control_tags = set(allowed_added_content_control_tags or set())

    try:
        with zipfile.ZipFile(target_docx, "r") as zf:
            bad = zf.testzip()
            if bad:
                issues.append(ValidationIssue("error", "zip_corrupt", f"Corrupt ZIP member: {bad}"))
    except Exception as exc:
        return ValidationReport(False, str(source_docx), str(target_docx), [ValidationIssue("error", "zip_open_failed", str(exc))], {})

    sh = _zip_hashes(source_docx)
    th = _zip_hashes(target_docx)
    source_parts = set(sh)
    target_parts = set(th)
    metrics["parts_added"] = sorted(target_parts - source_parts)
    metrics["parts_removed"] = sorted(source_parts - target_parts)
    if metrics["parts_removed"]:
        issues.append(ValidationIssue("error", "parts_removed", f"Removed ZIP parts: {metrics['parts_removed']}"))
    unexpected_added = [p for p in metrics["parts_added"] if p not in allowed_part_changes]
    if unexpected_added:
        issues.append(ValidationIssue("error" if strict else "warning", "parts_added", f"Added ZIP parts: {unexpected_added}"))

    changed = [p for p in sorted(source_parts & target_parts) if sh[p] != th[p]]
    metrics["changed_parts"] = changed
    unexpected_changes = [p for p in changed if p not in allowed_part_changes]
    if unexpected_changes:
        issues.append(ValidationIssue("error" if strict else "warning", "unexpected_part_change", f"Unexpected changed parts: {unexpected_changes}"))

    try:
        before = inspect_docx(source_docx)
        after = inspect_docx(target_docx)
        count_keys = ["table_count", "image_count", "field_count", "comment_count", "comment_reference_count", "tracked_change_count", "content_control_count", "heading_count"]
        for key in count_keys:
            if before[key] != after[key] and key not in allowed_count_changes:
                severity = "error" if strict else "warning"
                issues.append(ValidationIssue(severity, "structure_count_changed", f"{key} changed from {before[key]} to {after[key]}"))
        if before["paragraph_count"] != after["paragraph_count"] and not allow_paragraph_count_change and "paragraph_count" not in allowed_count_changes:
            issues.append(ValidationIssue("error" if strict else "warning", "paragraph_count_changed", f"paragraph_count changed from {before['paragraph_count']} to {after['paragraph_count']}"))
        before_tags = sorted(a["extra"].get("tag") for a in before["anchors"] if a["kind"] == "content_control")
        after_tags = sorted(a["extra"].get("tag") for a in after["anchors"] if a["kind"] == "content_control")
        after_tags_without_allowed_additions = sorted(t for t in after_tags if t not in allowed_added_content_control_tags)
        if before_tags != after_tags_without_allowed_additions:
            issues.append(ValidationIssue("error", "content_control_tags_changed", "Content control tags changed"))
        metrics["before"] = {k: before[k] for k in before if k.endswith("_count")}
        metrics["after"] = {k: after[k] for k in after if k.endswith("_count")}

        isolation_requested = bool(touched_content_control_tags or touched_para_ids or touched_paragraph_indices or touched_table_indices or touched_table_cells)
        if isolation_requested:
            sf = _structural_fingerprint(source_docx)
            tf = _structural_fingerprint(target_docx)
            protected_cc_changed = []
            for tag, h in sf["content_controls"].items():
                if tag in touched_content_control_tags:
                    continue
                if tf["content_controls"].get(tag) != h:
                    protected_cc_changed.append(tag)
            if protected_cc_changed:
                issues.append(ValidationIssue("error" if strict else "warning", "protected_content_control_changed", f"Untouched content controls changed: {protected_cc_changed[:20]}"))
            protected_tables_changed = []
            for idx, meta in sf["tables"].items():
                after_meta = tf["tables"].get(idx)
                if not after_meta:
                    continue
                if idx in touched_table_indices:
                    if not allow_table_dimension_change and (meta["row_count"] != after_meta["row_count"] or meta["column_counts"] != after_meta["column_counts"]):
                        issues.append(ValidationIssue("error", "touched_table_dimension_changed_without_permission", f"Table {idx} dimensions changed without permission"))
                    continue
                if after_meta["hash"] != meta["hash"]:
                    protected_tables_changed.append(idx)
            if protected_tables_changed:
                issues.append(ValidationIssue("error" if strict else "warning", "protected_table_changed", f"Untouched tables changed: {protected_tables_changed[:20]}"))

            protected_paras_changed = []
            for pid, meta in sf["paragraphs"].items():
                if pid in touched_para_ids:
                    continue
                if meta.get("content_control_tag") in touched_content_control_tags:
                    continue
                if meta.get("table_index") in touched_table_indices:
                    continue
                after_meta = tf["paragraphs"].get(pid)
                if after_meta and after_meta["hash"] != meta["hash"]:
                    protected_paras_changed.append(pid)
            if protected_paras_changed:
                issues.append(ValidationIssue("error" if strict else "warning", "protected_paragraph_changed", f"Untouched paragraphs with paraId changed: {protected_paras_changed[:20]}"))

            # Stronger guarantee for documents that do not carry w14:paraId:
            # every untouched top-level body block from the source must survive as an
            # unchanged, ordered subsequence in the target. This catches accidental
            # rewrites caused by paragraph-count-changing edits.
            seq_report = _sequence_preservation_report(
                sf.get("body_blocks", []),
                tf.get("body_blocks", []),
                touched_content_control_tags=touched_content_control_tags,
                touched_table_indices=touched_table_indices,
                touched_paragraph_indices=touched_paragraph_indices,
            )
            metrics["body_block_sequence"] = seq_report
            if not seq_report["ok"]:
                issues.append(ValidationIssue("error" if strict else "warning", "protected_body_block_changed", f"Untouched body blocks were modified/reordered/removed: {seq_report['missing_or_modified_protected_blocks'][:10]}"))

            protected_table_cells_changed = []
            for table_idx in touched_table_indices:
                before_cells = sf.get("table_cells", {}).get(table_idx)
                after_cells = tf.get("table_cells", {}).get(table_idx)
                if not before_cells or not after_cells:
                    continue
                if allow_table_dimension_change and after_cells.get("row_count", 0) < before_cells.get("row_count", 0):
                    issues.append(ValidationIssue("error", "touched_table_row_count_decreased", f"Table {table_idx} lost rows"))
                for cell_key, cell_hash in before_cells.get("cells", {}).items():
                    r, c = (int(x) for x in cell_key.split(":"))
                    if (table_idx, r, c) in touched_table_cells:
                        continue
                    if after_cells.get("cells", {}).get(cell_key) != cell_hash:
                        protected_table_cells_changed.append({"table_index": table_idx, "row": r, "col": c})
            if protected_table_cells_changed:
                issues.append(ValidationIssue("error" if strict else "warning", "protected_table_cell_changed", f"Untouched cells changed in touched tables: {protected_table_cells_changed[:20]}"))

            metrics["protected_object_checks"] = {
                "content_controls_checked": len(sf["content_controls"]),
                "paragraphs_with_para_id_checked": len(sf["paragraphs"]),
                "body_blocks_checked": len(sf.get("body_blocks", [])),
                "tables_checked": len(sf["tables"]),
                "table_cells_checked": sum(len(v.get("cells", {})) for v in sf.get("table_cells", {}).values()),
                "touched_content_control_tags": sorted(touched_content_control_tags),
                "touched_para_ids": sorted(touched_para_ids),
                "touched_paragraph_indices": sorted(touched_paragraph_indices),
                "touched_table_indices": sorted(touched_table_indices),
                "touched_table_cells": sorted(f"{t}:{r}:{c}" for t, r, c in touched_table_cells),
            }
        else:
            metrics["protected_object_checks"] = {"skipped": "no touched targets supplied; structural package checks only"}
    except Exception as exc:
        issues.append(ValidationIssue("error", "inspection_failed", str(exc)))

    ok = not any(i.severity == "error" for i in issues)
    return ValidationReport(ok, str(source_docx), str(target_docx), issues, metrics)


def compare_structure(source_docx: str | Path, target_docx: str | Path, strict: bool = True) -> dict[str, Any]:
    return validate_structure(source_docx, target_docx, strict=strict).to_dict()


def backup_docx(docx_path: str | Path, backup_dir: str | Path | None = None) -> str:
    docx_path = Path(docx_path)
    backup_dir = Path(backup_dir) if backup_dir else docx_path.parent / ".wordai" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup_path = backup_dir / f"{docx_path.stem}.{stamp}.{hashlib.sha1(docx_path.read_bytes()).hexdigest()[:8]}.docx"
    shutil.copy2(docx_path, backup_path)
    return str(backup_path)


def rollback_docx(backup_path: str | Path, restore_path: str | Path, make_backup_of_current: bool = True) -> dict[str, Any]:
    backup_path = Path(backup_path)
    restore_path = Path(restore_path)
    if not backup_path.exists():
        raise FileNotFoundError(str(backup_path))
    current_backup = None
    if restore_path.exists() and make_backup_of_current:
        current_backup = backup_docx(restore_path)
    restore_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(backup_path, restore_path)
    return {"restored_from": str(backup_path), "restored_to": str(restore_path), "backup_of_replaced_file": current_backup}


def write_sidecar_index(docx_path: str | Path, out_path: str | Path | None = None) -> str:
    data = inspect_docx(docx_path, include_text=False)
    data["outline"] = get_outline(docx_path)["headings"]
    data["tables"] = list_tables(docx_path)["tables"]
    data["health"] = health_check(docx_path)
    if out_path is None:
        out_path = Path(docx_path).parent / ".wordai" / "indexes" / f"{Path(docx_path).stem}.index.json"
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(out_path)

# ---------------------------------------------------------------------------
# Compatibility wrappers for the MCP server surface.  These keep the public tool
# names stable while the OOXML core stays intentionally conservative.
# ---------------------------------------------------------------------------

def extract_plain_text(docx_path: str | Path) -> str:
    return _extract_plain_text(docx_path)


def extract_comments(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, COMMENTS_XML):
        return {"comments": [], "comment_count": 0}
    tree = load_part_tree(docx_path, COMMENTS_XML)
    comments = []
    for c in tree.getroot().xpath("//w:comment", namespaces=NS):
        txt = element_text(c)
        comments.append(
            {
                "id": c.get(qn("w:id")),
                "author": c.get(qn("w:author")),
                "date": c.get(qn("w:date")),
                "text": txt,
                "text_sha256": _sha(txt),
            }
        )
    return {"comments": comments, "comment_count": len(comments)}


def part_hashes(docx_path: str | Path) -> dict[str, Any]:
    hashes = _zip_hashes(docx_path)
    return {"parts": [{"name": k, "sha256": v} for k, v in sorted(hashes.items())], "part_count": len(hashes)}


def restore_backup(backup_path: str | Path, target_path: str | Path) -> dict[str, Any]:
    return rollback_docx(backup_path, target_path)


def read_paragraph(docx_path: str | Path, paraId: str | None = None, paragraph_index: int | None = None, max_chars: int = 20000) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    p = _find_paragraph(root, para_id=paraId, paragraph_index=paragraph_index)
    if p is None:
        raise ValueError("Paragraph not found")
    text = paragraph_text(p)
    return {
        "paragraph_index": _para_index_map(root).get(p),
        "paraId": p.get(qn("w14:paraId")),
        "style_id": _style_id_for_paragraph(p),
        "text": text[:max_chars],
        "text_sha256": _sha(text),
        "truncated": len(text) > max_chars,
        "path": _element_path(p),
        "content_control_tag": _ancestor_sdt_tag(p),
        "table_index": _ancestor_table_index(p, root),
        "complexity": _complex_summary(p),
    }


def preflight_patchset(docx_path: str | Path, patchset: dict[str, Any]) -> dict[str, Any]:
    return dry_run_patchset(docx_path, patchset, keep_output=False)

# Compatibility wrappers for the Codex-friendly MCP surface.
def search_text(docx_path: str | Path, query: str, *, regex: bool = False, case_sensitive: bool = False, max_results: int = 50) -> dict[str, Any]:
    if regex:
        # Conservative regex support: use extracted paragraphs and return the same shape as find_text.
        pattern = re.compile(query, 0 if case_sensitive else re.I)
        tree = load_document_tree(docx_path)
        root = tree.getroot()
        results = []
        for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
            text = paragraph_text(p)
            if pattern.search(text):
                results.append({
                    "kind": "paragraph",
                    "paragraph_index": idx,
                    "paraId": p.get(qn("w14:paraId")),
                    "content_control_tag": _ancestor_sdt_tag(p),
                    "table_index": _ancestor_table_index(p, root),
                    "path": _element_path(p),
                    "context": text[:500],
                    "text_sha256": _sha(text),
                })
                if len(results) >= max_results:
                    break
        return {"query": query, "regex": True, "count": len(results), "results": results}
    return find_text(docx_path, query, case_sensitive=case_sensitive, max_results=max_results)


def extract_plain_text(docx_path: str | Path) -> str:
    return _extract_plain_text(docx_path)


def preflight_patchset(docx_path: str | Path, patchset: dict[str, Any]) -> dict[str, Any]:
    return dry_run_patchset(docx_path, patchset, keep_output=False)


def restore_backup(backup_path: str | Path, target_path: str | Path) -> dict[str, Any]:
    return rollback_docx(backup_path, target_path, make_backup_of_current=True)


def part_hashes(docx_path: str | Path) -> dict[str, Any]:
    return package_manifest(docx_path, include_hashes=True)


def read_paragraph(docx_path: str | Path, *, paraId: str | None = None, paragraph_index: int | None = None) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    p = _find_paragraph(root, para_id=paraId, paragraph_index=paragraph_index)
    if p is None:
        raise ValueError("paragraph not found")
    paras = root.xpath("//w:body//w:p", namespaces=NS)
    text = paragraph_text(p)
    return {
        "paragraph_index": paras.index(p) + 1,
        "paraId": p.get(qn("w14:paraId")),
        "style_id": _style_id_for_paragraph(p),
        "content_control_tag": _ancestor_sdt_tag(p),
        "table_index": _ancestor_table_index(p, root),
        "path": _element_path(p),
        "text": text,
        "text_sha256": _sha(text),
        "complexity": _complex_summary(p),
    }

# ---------------------------------------------------------------------------
# Codex-friendly aliases and sidecar exporters (v0.2 surface)
# ---------------------------------------------------------------------------


def list_paragraphs(docx_path: str | Path, max_preview: int = 240, include_empty: bool = False) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    styles = _load_paragraph_styles(docx_path)
    toc_paragraphs = _toc_paragraph_indices(root, styles)
    items: list[dict[str, Any]] = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        text = paragraph_text(p)
        if not include_empty and not text.strip():
            continue
        style_id = _style_id_for_paragraph(p)
        is_toc = idx in toc_paragraphs
        items.append(
            {
                "paragraph_index": idx,
                "paraId": p.get(qn("w14:paraId")),
                "style_id": style_id,
                "style_name": _style_name_for_id(styles, style_id),
                "heading_level": None if is_toc else _heading_level(style_id, styles, paragraph=p),
                "is_toc": is_toc,
                "content_control_tag": _ancestor_sdt_tag(p),
                "table_index": _ancestor_table_index(p, root),
                "path": _element_path(p),
                "text_preview": text[:max_preview],
                "text_sha256": _sha(text),
                "complexity": _complex_summary(p),
            }
        )
    return {"paragraphs": items, "count": len(items)}


def read_paragraph(docx_path: str | Path, paragraph_index: int | None = None, paraId: str | None = None) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    styles = _load_paragraph_styles(docx_path)
    toc_paragraphs = _toc_paragraph_indices(root, styles)
    p = _find_paragraph(root, para_id=paraId, paragraph_index=paragraph_index)
    if p is None:
        raise ValueError("Paragraph not found")
    text = paragraph_text(p)
    style_id = _style_id_for_paragraph(p)
    actual_index = _paragraph_global_index(root, p)
    is_toc = bool(actual_index and actual_index in toc_paragraphs)
    return {
        "paragraph_index": actual_index or paragraph_index,
        "paraId": p.get(qn("w14:paraId")),
        "style_id": style_id,
        "style_name": _style_name_for_id(styles, style_id),
        "heading_level": None if is_toc else _heading_level(style_id, styles, paragraph=p),
        "is_toc": is_toc,
        "content_control_tag": _ancestor_sdt_tag(p),
        "table_index": _ancestor_table_index(p, root),
        "path": _element_path(p),
        "text": text,
        "text_sha256": _sha(text),
        "complexity": _complex_summary(p),
    }


def list_headings(docx_path: str | Path, max_preview: int = 240) -> dict[str, Any]:
    outline = get_outline(docx_path)
    for h in outline["headings"]:
        if "text" in h:
            h["text_preview"] = h["text"][:max_preview]
            h["text_sha256"] = _sha(h["text"])
    return outline


def search_text(docx_path: str | Path, query: str, *, case_sensitive: bool = False, regex: bool = False, max_results: int = 50, context_chars: int = 120) -> dict[str, Any]:
    if not regex:
        return find_text(docx_path, query, case_sensitive=case_sensitive, max_results=max_results, context_chars=context_chars)
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    flags = 0 if case_sensitive else re.I
    pattern = re.compile(query, flags)
    results: list[dict[str, Any]] = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        text = paragraph_text(p)
        m = pattern.search(text)
        if not m:
            continue
        start = max(0, m.start() - context_chars)
        end = min(len(text), m.end() + context_chars)
        results.append(
            {
                "kind": "paragraph",
                "paragraph_index": idx,
                "paraId": p.get(qn("w14:paraId")),
                "content_control_tag": _ancestor_sdt_tag(p),
                "table_index": _ancestor_table_index(p, root),
                "path": _element_path(p),
                "context": text[start:end],
                "text_sha256": _sha(text),
            }
        )
        if len(results) >= max_results:
            break
    return {"query": query, "regex": True, "count": len(results), "results": results}


def list_fields(docx_path: str | Path, max_preview: int = 240) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    fields: list[dict[str, Any]] = []
    for idx, fld in enumerate(root.xpath("//w:fldSimple", namespaces=NS), start=1):
        fields.append({"field_index": idx, "kind": "fldSimple", "instruction": (fld.get(qn("w:instr")) or "")[:max_preview], "path": _element_path(fld)})
    base = len(fields)
    for j, instr in enumerate(root.xpath("//w:instrText", namespaces=NS), start=1):
        fields.append({"field_index": base + j, "kind": "complex", "instruction": (instr.text or "")[:max_preview], "path": _element_path(instr)})
    return {"fields": fields, "count": len(fields)}


def list_images(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    images: list[dict[str, Any]] = []
    for idx, node in enumerate(root.xpath("//w:drawing | //w:pict", namespaces=NS), start=1):
        rel_ids = []
        for el in node.iter():
            for k, v in el.attrib.items():
                if k.endswith("}embed") or k.endswith("}link"):
                    rel_ids.append(v)
        images.append({"image_index": idx, "relationship_ids": sorted(set(rel_ids)), "path": _element_path(node), "content_control_tag": _ancestor_sdt_tag(node), "table_index": _ancestor_table_index(node, root)})
    return {"images": images, "count": len(images)}


def list_comments(docx_path: str | Path, max_preview: int = 500) -> dict[str, Any]:
    if not _zip_has(docx_path, COMMENTS_XML):
        return {"comments": [], "count": 0}
    tree = _parse_xml(_zip_read(docx_path, COMMENTS_XML))
    comments: list[dict[str, Any]] = []
    for c in tree.getroot().xpath("//w:comment", namespaces=NS):
        cid = c.get(qn("w:id"))
        text = element_text(c)
        comments.append({"id": cid, "author": c.get(qn("w:author")), "date": c.get(qn("w:date")), "text_preview": text[:max_preview], "text_sha256": _sha(text)})
    return {"comments": comments, "count": len(comments)}


def list_revisions(docx_path: str | Path, max_preview: int = 240) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    revisions: list[dict[str, Any]] = []
    for idx, node in enumerate(tree.getroot().xpath("//w:ins | //w:del | //w:moveFrom | //w:moveTo", namespaces=NS), start=1):
        text = element_text(node)
        revisions.append({"revision_index": idx, "kind": etree.QName(node).localname, "author": node.get(qn("w:author")), "date": node.get(qn("w:date")), "path": _element_path(node), "text_preview": text[:max_preview], "text_sha256": _sha(text)})
    return {"revisions": revisions, "count": len(revisions)}


def plan_patchset(docx_path: str | Path, patchset: dict[str, Any]) -> dict[str, Any]:
    return assess_patchset(docx_path, patchset)


def export_table_csv(docx_path: str | Path, table_index: int, out_path: str | Path, scope_tag: str | None = None) -> str:
    table = read_table(docx_path, table_index, scope_tag=scope_tag, max_chars_per_cell=1_000_000)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        for row in table["rows"]:
            writer.writerow([cell["text"] for cell in row])
    return str(out_path)


# ---------------------------------------------------------------------------
# Additional v0.2 Codex/MCP compatibility helpers
# ---------------------------------------------------------------------------

def list_content_controls(docx_path: str | Path, max_preview: int = 500) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    controls: list[dict[str, Any]] = []
    for idx, sdt in enumerate(root.xpath("//w:sdt", namespaces=NS), start=1):
        text = get_content_control_text_from_tree(sdt)
        controls.append(
            {
                "index": idx,
                "tag": sdt_tag(sdt),
                "alias": sdt_alias(sdt),
                "id": sdt_id(sdt),
                "path": _element_path(sdt),
                "text_preview": text[:max_preview],
                "text_sha256": _sha(text),
                "char_count": len(text),
                "complexity": _complex_summary(sdt),
            }
        )
    return {"content_controls": controls, "count": len(controls)}


def read_table_cell(docx_path: str | Path, table_index: int, row: int, col: int, scope_tag: str | None = None) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    tables = _tables_in_scope(root, scope_tag)
    if table_index < 1 or table_index > len(tables):
        raise ValueError("table_index out of range")
    tc = _table_cell(tables[table_index - 1], row, col)
    text = _cell_text(tc)
    return {
        "table_index": table_index,
        "scope_tag": scope_tag,
        "row": row,
        "col": col,
        "text": text,
        "text_sha256": _sha(text),
        "path": _element_path(tc),
        "complexity": _complex_summary(tc),
    }


def list_hyperlinks(docx_path: str | Path, max_preview: int = 240) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    rel_targets: dict[str, dict[str, str]] = {}
    if _zip_has(docx_path, DOCUMENT_RELS):
        rel_tree = _parse_xml(_zip_read(docx_path, DOCUMENT_RELS))
        for rel in rel_tree.getroot():
            rid = rel.get("Id")
            if rid:
                rel_targets[rid] = {"type": rel.get("Type", ""), "target": rel.get("Target", ""), "target_mode": rel.get("TargetMode", "")}
    links: list[dict[str, Any]] = []
    for idx, link in enumerate(root.xpath("//w:hyperlink", namespaces=NS), start=1):
        rid = link.get(qn("r:id"))
        text = element_text(link) or "".join(t.text or "" for t in link.xpath(".//w:t", namespaces=NS))
        links.append(
            {
                "hyperlink_index": idx,
                "rId": rid,
                "anchor": link.get(qn("w:anchor")),
                "target": rel_targets.get(rid or "", {}).get("target"),
                "target_mode": rel_targets.get(rid or "", {}).get("target_mode"),
                "text_preview": text[:max_preview],
                "text_sha256": _sha(text),
                "path": _element_path(link),
            }
        )
    return {"hyperlinks": links, "count": len(links)}


def export_plain_text(docx_path: str | Path, out_path: str | Path | None = None) -> dict[str, Any]:
    text = _extract_plain_text(docx_path)
    if out_path is None:
        out_path = Path(docx_path).with_suffix(".txt")
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return {"out_path": str(out_path), "char_count": len(text), "text_sha256": _sha(text)}


def table_to_csv(docx_path: str | Path, table_index: int, out_path: str | Path | None = None, scope_tag: str | None = None) -> dict[str, Any]:
    if out_path is None:
        out_path = Path(docx_path).with_suffix(f".table{table_index}.csv")
    out = export_table_csv(docx_path, table_index, out_path, scope_tag=scope_tag)
    return {"out_path": out, "table_index": table_index, "scope_tag": scope_tag}


def validate_with_patchset(source_docx: str | Path, target_docx: str | Path, patchset: dict[str, Any], strict: bool = True) -> ValidationReport:
    assessment = assess_patchset(source_docx, patchset)
    allowed_part_changes = {DOCUMENT_XML}
    allowed_count_changes: set[str] = set()
    allow_table_dimension_change = False
    allow_paragraph_count_change = False
    for op in patchset.get("operations") or []:
        typ = op.get("op")
        if typ in {"append_content_control_text", "prepend_content_control_text", "insert_paragraph_after", "insert_paragraph_before"}:
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change = True
        if typ == "replace_content_control_text" and op.get("allow_paragraph_count_change"):
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change = True
        if typ == "replace_table_cell_text" and op.get("allow_paragraph_count_change"):
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change = True
        if typ == "append_table_row":
            allowed_count_changes.add("paragraph_count")
            allow_paragraph_count_change = True
            allow_table_dimension_change = True
        if typ == "wrap_paragraph_with_content_control":
            allowed_count_changes.add("content_control_count")
        if typ == "add_comment":
            allowed_part_changes.update({COMMENTS_XML, DOCUMENT_RELS, CONTENT_TYPES})
            allowed_count_changes.update({"comment_count", "comment_reference_count"})
    return validate_structure(
        source_docx,
        target_docx,
        strict=strict,
        allowed_part_changes=allowed_part_changes,
        allowed_count_changes=allowed_count_changes,
        touched_content_control_tags=assessment.get("touched", {}).get("content_control_tags", []),
        touched_para_ids=assessment.get("touched", {}).get("paraIds", []),
        touched_paragraph_indices=assessment.get("touched", {}).get("paragraph_indices", []),
        touched_table_indices=assessment.get("touched", {}).get("table_indices", []),
        touched_table_cells=assessment.get("touched", {}).get("table_cells", []),
        allow_table_dimension_change=allow_table_dimension_change,
        allow_paragraph_count_change=allow_paragraph_count_change,
        allowed_added_content_control_tags={str(op.get("tag")) for op in patchset.get("operations", []) if op.get("op") == "wrap_paragraph_with_content_control" and op.get("tag")},
    )

# ---------------------------------------------------------------------------
# Final stable public wrappers used by word_ai_mcp.server
# ---------------------------------------------------------------------------

def load_part_tree(docx_path: str | Path, part_name: str) -> etree._ElementTree:
    return _parse_xml(_zip_read(docx_path, part_name))


def list_docx_parts(docx_path: str | Path) -> dict[str, Any]:
    return package_manifest(docx_path, include_hashes=True)


def build_document_map(docx_path: str | Path, max_preview: int = 500, include_text: bool = False) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    styles = _load_paragraph_styles(docx_path)
    toc_paragraphs = _toc_paragraph_indices(root, styles)
    paragraphs = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        text = paragraph_text(p)
        if not text.strip() and not include_text:
            continue
        style_id = _style_id_for_paragraph(p)
        is_toc = idx in toc_paragraphs
        paragraphs.append(
            {
                "paragraph_index": idx,
                "paraId": p.get(qn("w14:paraId")),
                "style_id": style_id,
                "style_name": _style_name_for_id(styles, style_id),
                "heading_level": None if is_toc else _heading_level(style_id, styles, paragraph=p),
                "is_toc": is_toc,
                "content_control_tag": _ancestor_sdt_tag(p),
                "table_index": _ancestor_table_index(p, root),
                "text_preview": text[:max_preview],
                "text_sha256": _sha(text),
                "path": _element_path(p),
                "complexity": _complex_summary(p),
            }
        )
    return {
        "path": str(docx_path),
        "health": health_check(docx_path),
        "headings": get_outline(docx_path)["headings"],
        "content_controls": [a.to_dict() for a in list_anchors(docx_path, max_preview) if a.kind == "content_control"],
        "paragraphs": paragraphs,
        "tables": list_tables(docx_path, max_cell_chars=min(max_preview, 500))["tables"],
    }


def list_styles(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, STYLES_XML):
        return {"styles": [], "count": 0}
    tree = _parse_xml(_zip_read(docx_path, STYLES_XML))
    styles = []
    for style in tree.getroot().xpath("//w:style", namespaces=NS):
        name = style.find("./w:name", namespaces=NS)
        based_on = style.find("./w:basedOn", namespaces=NS)
        styles.append(
            {
                "style_id": style.get(qn("w:styleId")),
                "type": style.get(qn("w:type")),
                "default": style.get(qn("w:default")),
                "name": name.get(qn("w:val")) if name is not None else None,
                "based_on": based_on.get(qn("w:val")) if based_on is not None else None,
            }
        )
    return {"styles": styles, "count": len(styles)}


def list_numbering(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, NUMBERING_XML):
        return {"abstractNums": [], "nums": [], "count": 0}
    tree = _parse_xml(_zip_read(docx_path, NUMBERING_XML))
    root = tree.getroot()
    abstract_nums = []
    for abs_num in root.xpath("//w:abstractNum", namespaces=NS):
        levels = []
        for lvl in abs_num.xpath("./w:lvl", namespaces=NS):
            num_fmt = lvl.find("./w:numFmt", namespaces=NS)
            lvl_text = lvl.find("./w:lvlText", namespaces=NS)
            levels.append({"ilvl": lvl.get(qn("w:ilvl")), "numFmt": num_fmt.get(qn("w:val")) if num_fmt is not None else None, "lvlText": lvl_text.get(qn("w:val")) if lvl_text is not None else None})
        abstract_nums.append({"abstractNumId": abs_num.get(qn("w:abstractNumId")), "levels": levels})
    nums = []
    for num in root.xpath("//w:num", namespaces=NS):
        abs_id = num.find("./w:abstractNumId", namespaces=NS)
        nums.append({"numId": num.get(qn("w:numId")), "abstractNumId": abs_id.get(qn("w:val")) if abs_id is not None else None})
    return {"abstractNums": abstract_nums, "nums": nums, "count": len(nums)}


def _twips_to_mm(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return round(int(value) * 25.4 / 1440, 2)
    except Exception:
        return None


def list_sections(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    sections = []
    for idx, sect in enumerate(tree.getroot().xpath("//w:sectPr", namespaces=NS), start=1):
        pg_sz = sect.find("./w:pgSz", namespaces=NS)
        pg_mar = sect.find("./w:pgMar", namespaces=NS)
        refs = []
        for ref in sect.xpath("./w:headerReference | ./w:footerReference", namespaces=NS):
            refs.append({"kind": etree.QName(ref).localname, "type": ref.get(qn("w:type")), "rId": ref.get(qn("r:id"))})
        sections.append(
            {
                "section_index": idx,
                "page_size_twips": {"w": pg_sz.get(qn("w:w")) if pg_sz is not None else None, "h": pg_sz.get(qn("w:h")) if pg_sz is not None else None, "orient": pg_sz.get(qn("w:orient")) if pg_sz is not None else None},
                "margins_mm": {
                    "top": _twips_to_mm(pg_mar.get(qn("w:top")) if pg_mar is not None else None),
                    "bottom": _twips_to_mm(pg_mar.get(qn("w:bottom")) if pg_mar is not None else None),
                    "left": _twips_to_mm(pg_mar.get(qn("w:left")) if pg_mar is not None else None),
                    "right": _twips_to_mm(pg_mar.get(qn("w:right")) if pg_mar is not None else None),
                },
                "header_footer_refs": refs,
                "path": _element_path(sect),
            }
        )
    return {"sections": sections, "count": len(sections)}

# ---------------------------------------------------------------------------
# v0.4 stable public helpers used by the richer MCP tool surface.
# ---------------------------------------------------------------------------

def load_part_tree(docx_path: str | Path, member: str) -> etree._ElementTree:
    return _parse_xml(_zip_read(docx_path, member))


def get_part_manifest(docx_path: str | Path, include_hashes: bool = True) -> dict[str, Any]:
    return package_manifest(docx_path, include_hashes=include_hashes)


def structural_fingerprint(docx_path: str | Path) -> dict[str, Any]:
    fp = _structural_fingerprint(docx_path)
    return {
        "path": str(docx_path),
        "package": package_manifest(docx_path, include_hashes=True),
        "content_controls": fp.get("content_controls", {}),
        "paragraphs": fp.get("paragraphs", {}),
        "tables": fp.get("tables", {}),
        "outline": get_outline(docx_path),
        "health": health_check(docx_path),
    }


def list_content_controls(docx_path: str | Path, max_preview: int = 240) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    controls: list[dict[str, Any]] = []
    for idx, sdt in enumerate(tree.getroot().xpath("//w:sdt", namespaces=NS), start=1):
        text = get_content_control_text_from_tree(sdt)
        controls.append(
            {
                "index": idx,
                "tag": sdt_tag(sdt),
                "alias": sdt_alias(sdt),
                "id": sdt_id(sdt),
                "path": _element_path(sdt),
                "text_preview": text[:max_preview],
                "text_sha256": _sha(text),
                "xml_sha256": _canonical_hash(sdt),
                "complexity": _complex_summary(sdt),
            }
        )
    return {"content_controls": controls, "count": len(controls)}


def read_table_cell(docx_path: str | Path, table_index: int, row: int, col: int, scope_tag: str | None = None, max_chars: int = 20000) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    tables = _tables_in_scope(root, scope_tag)
    if table_index < 1 or table_index > len(tables):
        raise ValueError("table_index out of range")
    cell = _table_cell(tables[table_index - 1], row, col)
    text = _cell_text(cell)
    return {
        "table_index": table_index,
        "scope_tag": scope_tag,
        "row": row,
        "col": col,
        "text": text[:max_chars],
        "text_sha256": _sha(text),
        "truncated": len(text) > max_chars,
        "path": _element_path(cell),
        "complexity": _complex_summary(cell),
    }


def list_bookmarks(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    bookmarks: list[dict[str, Any]] = []
    for bm in tree.getroot().xpath("//w:bookmarkStart", namespaces=NS):
        name = bm.get(qn("w:name"))
        if name:
            bookmarks.append({"name": name, "id": bm.get(qn("w:id")), "path": _element_path(bm)})
    return {"bookmarks": bookmarks, "count": len(bookmarks)}


def list_headers_footers(docx_path: str | Path, max_preview: int = 500) -> dict[str, Any]:
    parts: list[dict[str, Any]] = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in sorted(zf.namelist()):
            if not re.fullmatch(r"word/(header|footer)\d+\.xml", name):
                continue
            tree = _parse_xml(zf.read(name))
            text = "\n".join(paragraph_text(p) for p in tree.getroot().xpath("//w:p", namespaces=NS))
            parts.append({"part": name, "text_preview": text[:max_preview], "text_sha256": _sha(text), "xml_sha256": _canonical_hash(tree.getroot())})
    return {"headers_footers": parts, "count": len(parts)}


def list_notes(docx_path: str | Path, max_preview: int = 500) -> dict[str, Any]:
    notes: list[dict[str, Any]] = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        names = set(zf.namelist())
        for part, tag in [("word/footnotes.xml", "footnote"), ("word/endnotes.xml", "endnote")]:
            if part not in names:
                continue
            tree = _parse_xml(zf.read(part))
            for node in tree.getroot().xpath(f"//w:{tag}", namespaces=NS):
                text = element_text(node)
                notes.append({"part": part, "id": node.get(qn("w:id")), "text_preview": text[:max_preview], "text_sha256": _sha(text)})
    return {"notes": notes, "count": len(notes)}


def preview_patchset(docx_path: str | Path, patchset: dict[str, Any]) -> dict[str, Any]:
    return assess_patchset(docx_path, patchset)

# ---------------------------------------------------------------------------
# Final MCP surface shims.  These definitions are intentionally placed last so
# they override older narrow helpers while reusing the conservative core above.
# ---------------------------------------------------------------------------

_list_anchors_core = list_anchors
_read_anchor_core = read_anchor
_read_table_core = read_table


def list_docx_parts(docx_path: str | Path) -> dict[str, Any]:
    return package_manifest(docx_path, include_hashes=True)


def list_anchors(docx_path: str | Path, max_preview: int = 240, include_table_cells: bool = True) -> list[Anchor]:
    anchors = _list_anchors_core(docx_path, max_preview)
    if not include_table_cells:
        return anchors
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    tables = root.xpath("//w:body//w:tbl", namespaces=NS)
    for ti, tbl in enumerate(tables, start=1):
        for ri, tr in enumerate(tbl.xpath("./w:tr", namespaces=NS), start=1):
            for ci, tc in enumerate(tr.xpath("./w:tc", namespaces=NS), start=1):
                txt = _cell_text(tc)
                anchors.append(
                    Anchor(
                        anchor_id=f"table:{ti}:r{ri}:c{ci}",
                        kind="table_cell",
                        label=f"Table {ti} R{ri}C{ci}",
                        path=_element_path(tc),
                        text_preview=txt[:max_preview],
                        extra={"table_index": ti, "row": ri, "col": ci, "text_sha256": _sha(txt), "content_control_tag": _ancestor_sdt_tag(tc)},
                    )
                )
    return anchors


def read_anchor(docx_path: str | Path, anchor: str | dict[str, Any], max_chars: int = 20000) -> dict[str, Any]:
    if isinstance(anchor, str):
        return _read_anchor_core(docx_path, anchor, max_chars=max_chars)
    if not isinstance(anchor, dict):
        raise ValueError("anchor must be string or object")
    if anchor.get("anchor_id"):
        return _read_anchor_core(docx_path, str(anchor["anchor_id"]), max_chars=max_chars)
    if anchor.get("tag"):
        data = get_content_control_text(docx_path, str(anchor["tag"]))
        data["text"] = data["text"][:max_chars]
        return data
    if anchor.get("paraId") or anchor.get("paragraph_index"):
        return read_paragraph(docx_path, paraId=anchor.get("paraId"), paragraph_index=anchor.get("paragraph_index"))
    if anchor.get("table_index") and anchor.get("row") and anchor.get("col"):
        table_index = int(anchor["table_index"])
        row = int(anchor["row"])
        col = int(anchor["col"])
        tree = load_document_tree(docx_path)
        root = tree.getroot()
        tables = root.xpath("//w:body//w:tbl", namespaces=NS)
        if table_index < 1 or table_index > len(tables):
            raise ValueError("table_index out of range")
        tc = _table_cell(tables[table_index - 1], row, col)
        txt = _cell_text(tc)
        return {"kind": "table_cell", "table_index": table_index, "row": row, "col": col, "text": txt[:max_chars], "text_sha256": _sha(txt), "path": _element_path(tc)}
    raise ValueError("Unsupported anchor object")


def read_table(docx_path: str | Path, table_index: int, max_rows: int | None = None, scope_tag: str | None = None, max_chars_per_cell: int = 2000) -> dict[str, Any]:
    data = _read_table_core(docx_path, table_index, scope_tag=scope_tag, max_chars_per_cell=max_chars_per_cell)
    if max_rows is not None:
        data["rows"] = data["rows"][: int(max_rows)]
        data["truncated"] = True
    return data


def build_document_map(docx_path: str | Path, max_preview: int = 500, include_text: bool = False) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    styles = _load_paragraph_styles(docx_path)
    toc_paragraphs = _toc_paragraph_indices(root, styles)
    paragraphs = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        txt = paragraph_text(p)
        style_id = _style_id_for_paragraph(p)
        is_toc = idx in toc_paragraphs
        item = {
            "paragraph_index": idx,
            "paraId": p.get(qn("w14:paraId")),
            "style_id": style_id,
            "style_name": _style_name_for_id(styles, style_id),
            "heading_level": None if is_toc else _heading_level(style_id, styles, paragraph=p),
            "is_toc": is_toc,
            "content_control_tag": _ancestor_sdt_tag(p),
            "table_index": _ancestor_table_index(p, root),
            "text_preview": txt[:max_preview],
            "text_sha256": _sha(txt),
            "path": _element_path(p),
        }
        if include_text:
            item["text"] = txt
        paragraphs.append(item)
    return {
        "document": {"path": str(docx_path), "sha256": hashlib.sha256(Path(docx_path).read_bytes()).hexdigest()},
        "outline": get_outline(docx_path)["headings"],
        "paragraphs": paragraphs,
        "tables": list_tables(docx_path)["tables"],
        "anchors_count": len(list_anchors(docx_path, max_preview=max_preview, include_table_cells=False)),
    }


def list_styles(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, STYLES_XML):
        return {"styles": [], "count": 0}
    tree = load_part_tree(docx_path, STYLES_XML)
    styles = []
    for s in tree.getroot().xpath("//w:style", namespaces=NS):
        name = s.find("./w:name", namespaces=NS)
        styles.append({"style_id": s.get(qn("w:styleId")), "type": s.get(qn("w:type")), "name": name.get(qn("w:val")) if name is not None else None})
    return {"styles": styles, "count": len(styles)}


def list_numbering(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, NUMBERING_XML):
        return {"abstract_nums": [], "nums": [], "abstract_num_count": 0, "num_count": 0}
    tree = load_part_tree(docx_path, NUMBERING_XML)
    root = tree.getroot()
    abstract_nums = [a.get(qn("w:abstractNumId")) for a in root.xpath("//w:abstractNum", namespaces=NS)]
    nums = []
    for n in root.xpath("//w:num", namespaces=NS):
        absn = n.find("./w:abstractNumId", namespaces=NS)
        nums.append({"numId": n.get(qn("w:numId")), "abstractNumId": absn.get(qn("w:val")) if absn is not None else None})
    return {"abstract_nums": abstract_nums, "nums": nums, "abstract_num_count": len(abstract_nums), "num_count": len(nums)}


def list_sections(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    sections = []
    for idx, sect in enumerate(tree.getroot().xpath("//w:sectPr", namespaces=NS), start=1):
        pg_sz = sect.find("./w:pgSz", namespaces=NS)
        pg_mar = sect.find("./w:pgMar", namespaces=NS)
        headers = [h.get(qn("r:id")) for h in sect.xpath("./w:headerReference", namespaces=NS)]
        footers = [f.get(qn("r:id")) for f in sect.xpath("./w:footerReference", namespaces=NS)]
        sections.append(
            {
                "section_index": idx,
                "path": _element_path(sect),
                "page_size": dict(pg_sz.attrib) if pg_sz is not None else {},
                "page_margin": dict(pg_mar.attrib) if pg_mar is not None else {},
                "header_relationship_ids": headers,
                "footer_relationship_ids": footers,
            }
        )
    return {"sections": sections, "count": len(sections)}

# --- Additional read-only / sidecar utilities for broader Word scenarios ---

def list_docx_parts(docx_path: str | Path) -> dict[str, Any]:
    return package_manifest(docx_path, include_hashes=True)


def build_document_map(docx_path: str | Path, max_preview: int = 500, include_text: bool = False) -> dict[str, Any]:
    return {
        "inspect": inspect_docx(docx_path, include_text=include_text, max_preview=max_preview),
        "outline": get_outline(docx_path),
        "tables": list_tables(docx_path),
        "health": health_check(docx_path),
    }


def list_styles(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, STYLES_XML):
        return {"styles": [], "style_count": 0}
    tree = _parse_xml(_zip_read(docx_path, STYLES_XML))
    paragraph_styles = _load_paragraph_styles(docx_path)
    styles = []
    for st in tree.getroot().xpath("//w:style", namespaces=NS):
        name = st.find("./w:name", namespaces=NS)
        based = st.find("./w:basedOn", namespaces=NS)
        nxt = st.find("./w:next", namespaces=NS)
        outline = st.find("./w:pPr/w:outlineLvl", namespaces=NS)
        style_id = st.get(qn("w:styleId"))
        paragraph_info = paragraph_styles.get(style_id or "")
        styles.append({
            "style_id": style_id,
            "type": st.get(qn("w:type")),
            "default": st.get(qn("w:default")),
            "name": name.get(qn("w:val")) if name is not None else None,
            "based_on": based.get(qn("w:val")) if based is not None else None,
            "next": nxt.get(qn("w:val")) if nxt is not None else None,
            "outline_level": outline.get(qn("w:val")) if outline is not None else None,
            "heading_level": paragraph_info.get("heading_level") if paragraph_info else None,
            "is_toc": bool(paragraph_info.get("is_toc")) if paragraph_info else _is_toc_style(style_id),
        })
    return {"style_count": len(styles), "styles": styles}


def list_numbering(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, NUMBERING_XML):
        return {"abstract_num_count": 0, "num_count": 0, "nums": []}
    tree = _parse_xml(_zip_read(docx_path, NUMBERING_XML))
    root = tree.getroot()
    nums = []
    for num in root.xpath("//w:num", namespaces=NS):
        aid = num.find("./w:abstractNumId", namespaces=NS)
        nums.append({"numId": num.get(qn("w:numId")), "abstractNumId": aid.get(qn("w:val")) if aid is not None else None})
    return {
        "abstract_num_count": len(root.xpath("//w:abstractNum", namespaces=NS)),
        "num_count": len(nums),
        "nums": nums,
    }


def list_sections(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    sections = []
    for idx, sect in enumerate(root.xpath("//w:sectPr", namespaces=NS), start=1):
        pg_sz = sect.find("./w:pgSz", namespaces=NS)
        pg_mar = sect.find("./w:pgMar", namespaces=NS)
        refs = []
        for ref in sect.xpath("./w:headerReference | ./w:footerReference", namespaces=NS):
            refs.append({"kind": etree.QName(ref).localname, "type": ref.get(qn("w:type")), "rId": ref.get(qn("r:id"))})
        sections.append({
            "section_index": idx,
            "path": _element_path(sect),
            "page_size": dict(pg_sz.attrib) if pg_sz is not None else {},
            "page_margins": dict(pg_mar.attrib) if pg_mar is not None else {},
            "header_footer_refs": refs,
        })
    return {"section_count": len(sections), "sections": sections}


def list_fields(docx_path: str | Path, max_instr: int = 200) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    fields = []
    for el in root.xpath("//w:fldSimple", namespaces=NS):
        instr = el.get(qn("w:instr")) or ""
        fields.append({"kind": "fldSimple", "instr": instr[:max_instr], "path": _element_path(el)})
    for el in root.xpath("//w:instrText", namespaces=NS):
        instr = el.text or ""
        fields.append({"kind": "instrText", "instr": instr[:max_instr], "path": _element_path(el)})
    return {"field_count": len(fields), "fields": fields}


def list_images(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    images = []
    for idx, drawing in enumerate(root.xpath("//w:drawing | //w:pict", namespaces=NS), start=1):
        docpr = drawing.find(".//wp:docPr", namespaces=NS)
        blips = drawing.xpath(".//*[local-name()='blip']")
        images.append({
            "image_index": idx,
            "path": _element_path(drawing),
            "content_control_tag": _ancestor_sdt_tag(drawing),
            "docPr_id": docpr.get("id") if docpr is not None else None,
            "name": docpr.get("name") if docpr is not None else None,
            "title": docpr.get("title") if docpr is not None else None,
            "descr": docpr.get("descr") if docpr is not None else None,
            "relationship_ids": [b.get(qn("r:embed")) or b.get(qn("r:link")) for b in blips],
        })
    return {"image_count": len(images), "images": images}


def list_comments(docx_path: str | Path) -> dict[str, Any]:
    if not _zip_has(docx_path, COMMENTS_XML):
        return {"comment_count": 0, "comments": []}
    tree = _parse_xml(_zip_read(docx_path, COMMENTS_XML))
    comments = []
    for c in tree.getroot().xpath("//w:comment", namespaces=NS):
        text = "\n".join(paragraph_text(p) for p in c.xpath(".//w:p", namespaces=NS))
        comments.append({
            "id": c.get(qn("w:id")),
            "author": c.get(qn("w:author")),
            "date": c.get(qn("w:date")),
            "initials": c.get(qn("w:initials")),
            "text": text,
            "text_sha256": _sha(text),
        })
    return {"comment_count": len(comments), "comments": comments}


def export_table_csv(docx_path: str | Path, table_index: int, out_path: str | Path, scope_tag: str | None = None) -> str:
    import csv
    table = read_table(docx_path, table_index, scope_tag=scope_tag)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        for row in table["rows"]:
            writer.writerow([cell["text"] for cell in row])
    return str(out_path)


# Backward-compatible alias used by older docs/scripts.
def preflight_patchset(docx_path: str | Path, patchset: dict[str, Any]) -> dict[str, Any]:
    return assess_patchset(docx_path, patchset)


# Backward-compatible alias.
def search_text(docx_path: str | Path, query: str, case_sensitive: bool = False, max_results: int = 50) -> dict[str, Any]:
    return find_text(docx_path, query, case_sensitive=case_sensitive, max_results=max_results)

# Final overriding shims for the current server signatures.
def list_comments(docx_path: str | Path, max_preview: int = 500) -> dict[str, Any]:
    if not _zip_has(docx_path, COMMENTS_XML):
        return {"comment_count": 0, "comments": []}
    tree = _parse_xml(_zip_read(docx_path, COMMENTS_XML))
    comments = []
    for c in tree.getroot().xpath("//w:comment", namespaces=NS):
        text = "\n".join(paragraph_text(p) for p in c.xpath(".//w:p", namespaces=NS))
        comments.append({
            "id": c.get(qn("w:id")),
            "author": c.get(qn("w:author")),
            "date": c.get(qn("w:date")),
            "initials": c.get(qn("w:initials")),
            "text": text[:max_preview],
            "truncated": len(text) > max_preview,
            "text_sha256": _sha(text),
        })
    return {"comment_count": len(comments), "comments": comments}


def search_text(docx_path: str | Path, query: str, *, regex: bool = False, case_sensitive: bool = False, max_results: int = 50, context_chars: int = 120) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    flags = 0 if case_sensitive else re.I
    pattern = re.compile(query if regex else re.escape(query), flags)
    results: list[dict[str, Any]] = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        text = paragraph_text(p)
        m = pattern.search(text)
        if not m:
            continue
        start = max(0, m.start() - context_chars)
        end = min(len(text), m.end() + context_chars)
        results.append({
            "kind": "paragraph",
            "paragraph_index": idx,
            "paraId": p.get(qn("w14:paraId")),
            "style_id": _style_id_for_paragraph(p),
            "content_control_tag": _ancestor_sdt_tag(p),
            "table_index": _ancestor_table_index(p, root),
            "path": _element_path(p),
            "match": m.group(0),
            "context": text[start:end],
            "text_sha256": _sha(text),
        })
        if len(results) >= max_results:
            break
    return {"query": query, "regex": regex, "case_sensitive": case_sensitive, "count": len(results), "results": results}

# --- Final override shims for server signatures ---
def search_text(docx_path: str | Path, query: str, regex: bool = False, case_sensitive: bool = False, max_results: int = 50, context_chars: int = 120) -> dict[str, Any]:
    if not regex:
        return find_text(docx_path, query, case_sensitive=case_sensitive, max_results=max_results, context_chars=context_chars)
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    flags = 0 if case_sensitive else re.I
    pattern = re.compile(query, flags)
    results: list[dict[str, Any]] = []
    for idx, p in enumerate(root.xpath("//w:body//w:p", namespaces=NS), start=1):
        text = paragraph_text(p)
        m = pattern.search(text)
        if not m:
            continue
        start = max(0, m.start() - context_chars)
        end = min(len(text), m.end() + context_chars)
        results.append({
            "kind": "paragraph",
            "paragraph_index": idx,
            "paraId": p.get(qn("w14:paraId")),
            "content_control_tag": _ancestor_sdt_tag(p),
            "table_index": _ancestor_table_index(p, root),
            "path": _element_path(p),
            "context": text[start:end],
            "text_sha256": _sha(text),
        })
        if len(results) >= max_results:
            break
    return {"query": query, "regex": True, "count": len(results), "results": results}


def list_comments(docx_path: str | Path, max_preview: int = 500) -> dict[str, Any]:
    if not _zip_has(docx_path, COMMENTS_XML):
        return {"comment_count": 0, "comments": []}
    tree = _parse_xml(_zip_read(docx_path, COMMENTS_XML))
    comments = []
    for c in tree.getroot().xpath("//w:comment", namespaces=NS):
        text = "\n".join(paragraph_text(p) for p in c.xpath(".//w:p", namespaces=NS))
        comments.append({
            "id": c.get(qn("w:id")),
            "author": c.get(qn("w:author")),
            "date": c.get(qn("w:date")),
            "initials": c.get(qn("w:initials")),
            "text_preview": text[:max_preview],
            "text_sha256": _sha(text),
        })
    return {"comment_count": len(comments), "comments": comments}

# Server-compatible final table reader signature: positional argument #3 is scope_tag,
# positional argument #4 is max_chars_per_cell.
def read_table(docx_path: str | Path, table_index: int, scope_tag: str | None = None, max_chars_per_cell: int = 2000, max_rows: int | None = None) -> dict[str, Any]:
    data = _read_table_core(docx_path, table_index, scope_tag=scope_tag, max_chars_per_cell=max_chars_per_cell)
    if max_rows is not None:
        data["rows"] = data["rows"][: int(max_rows)]
        data["truncated"] = True
    return data

# Count-key compatibility for smoke tests and generic clients.
def list_fields(docx_path: str | Path, max_preview: int = 240, max_instr: int | None = None) -> dict[str, Any]:
    limit = max_instr if max_instr is not None else max_preview
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    fields = []
    for el in root.xpath("//w:fldSimple", namespaces=NS):
        instr = el.get(qn("w:instr")) or ""
        fields.append({"kind": "fldSimple", "instr": instr[:limit], "path": _element_path(el)})
    for el in root.xpath("//w:instrText", namespaces=NS):
        instr = el.text or ""
        fields.append({"kind": "instrText", "instr": instr[:limit], "path": _element_path(el)})
    return {"count": len(fields), "field_count": len(fields), "fields": fields}


def list_images(docx_path: str | Path) -> dict[str, Any]:
    tree = load_document_tree(docx_path)
    root = tree.getroot()
    images = []
    for idx, drawing in enumerate(root.xpath("//w:drawing | //w:pict", namespaces=NS), start=1):
        docpr = drawing.find(".//wp:docPr", namespaces=NS)
        blips = drawing.xpath(".//*[local-name()='blip']")
        images.append({
            "image_index": idx,
            "path": _element_path(drawing),
            "content_control_tag": _ancestor_sdt_tag(drawing),
            "docPr_id": docpr.get("id") if docpr is not None else None,
            "name": docpr.get("name") if docpr is not None else None,
            "title": docpr.get("title") if docpr is not None else None,
            "descr": docpr.get("descr") if docpr is not None else None,
            "relationship_ids": [b.get(qn("r:embed")) or b.get(qn("r:link")) for b in blips],
        })
    return {"count": len(images), "image_count": len(images), "images": images}
