"""Conservative front-matter identities shared by profiles and target selection.

These identities supplement the historical module/role fields. They never
provide default formatting or change document text.
"""

import re

from document_structure import (
    MODULE_MARKERS,
    abstract_label_ranges,
    is_inline_abstract,
    keyword_label_ranges,
)

from .namespaces import qn
from .properties import read_paragraph

_METADATA_LABEL_RE = re.compile(
    r"^(?:(?:论文|设计)?题目|(?:学生|作者|研究生)?姓名|学号|(?:所属|二级|所在)?学院|"
    r"院[（(]?系[）)]?|系别|(?:所学|学科)?专业|研究方向|(?:专业|行政)?班级|年级|"
    r"(?:校内|企业|校外)?(?:指导教师|指导老师|导师)|职称|日期|时间|提交日期|完成日期|"
    r"分类号|中图分类号|密级|udc|学校代码|单位代码|"
    r"(?:本人|作者|我们)?(?:郑重|特此|谨此)?声明)[:：]",
    re.I,
)


def language(text):
    if re.search(r"英文|外文", text):
        return "en"
    return "zh" if re.search(r"[\u4e00-\u9fff]", text) else "en"


def _has_page_or_section_break(elem) -> bool:
    if elem.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
        return True
    return any(
        br.get(qn("w:type")) == "page"
        and not any(a.tag == qn("w:txbxContent") for a in br.iterancestors())
        for br in elem.iter(qn("w:br"))
    )


def _has_separate_cover(index, top_ids: list[str], first_abs_idx: int) -> bool:
    cover_meta_roles = {
        "cover_subject",
        "cover_student",
        "cover_college",
        "cover_major",
        "cover_advisor",
        "cover_student_id",
        "cover_date",
    }
    for rec in index.records.values():
        if rec.get("semantic_role") in cover_meta_roles or rec.get("role") == "cover_text":
            return True
    for oid in top_ids[:first_abs_idx]:
        rec = index.records[oid]
        elem = index.elements[oid]
        if rec.get("module") == "declaration":
            return True
        if re.search(r"封面|毕业(?:设计|论文)|学位论文", rec.get("text", "")):
            return True
        if _has_page_or_section_break(elem):
            return True
        if any(n.tag in {qn("w:drawing"), qn("w:pict")} for n in elem.iter()):
            return True
    if 0 <= first_abs_idx < len(top_ids):
        abs_elem = index.elements[top_ids[first_abs_idx]]
        if bool(read_paragraph(index.document, abs_elem, "page_break_before")):
            return True
    return False


def _is_abstract_title_candidate(record, *, allow_cover_title: bool = False) -> bool:
    text = record.get("text", "").strip()
    if not text or len(text) > 200:
        return False
    compact = re.sub(r"\s+", "", text)
    allowed_semantics = {None, "cover_title"} if allow_cover_title else {None}
    return (
        record.get("role") in {"body", "module_body", "title"}
        and record.get("module") in {None, "cover", "declaration", "abstract", "keywords"}
        and record.get("semantic_role") in allowed_semantics
        and not text.endswith((":", "："))
        and not re.search(r"[。；;！？!?]|签名|签字|年.*月", text)
        and not _METADATA_LABEL_RE.match(compact)
        and not any(p.fullmatch(text) for p in MODULE_MARKERS.values())
    )


def annotate(index):
    records = [index.records[oid] for oid in index.top_ids]
    for record in index.records.values():
        if record.get("kind") != "paragraph":
            continue
        text = record.get("text", "").strip()
        compact = re.sub(r"\s+", "", text)
        if record.get("module") == "cover":
            for pattern, identity in (
                (r"^(?:论文|设计)?题目[:：]", "cover_subject"),
                (r"^(?:学生|作者)?姓名[:：]", "cover_student"),
                (r"^(?:所属|二级|所在)?(?:学院|院[（(]?系[）)]?|系别)[:：]", "cover_college"),
                (r"^(?:所学|学科)?专业[:：]", "cover_major"),
                (r"^(?:校内)?(?:指导教师|指导老师|导师)[:：]", "cover_advisor"),
                (r"^(?:学号|学生学号)[:：]", "cover_student_id"),
                (r"^(?:日期[:：])?[\d〇○零一二三四五六七八九十]{4}年.*月", "cover_date"),
                (r"^(?:提交|完成)?日期[:：]", "cover_date"),
            ):
                if re.search(pattern, compact):
                    record["semantic_role"] = identity
                    break
            else:
                if record.get("role") == "title":
                    record["semantic_role"] = "cover_title"
        if MODULE_MARKERS["abstract"].fullmatch(text):
            if is_inline_abstract(text):
                ranges = abstract_label_ranges(text)
                lbl = text[ranges["label"][0] : ranges["label"][1]] if ranges else text
                record.update(semantic_role="abstract_inline", language=language(lbl))
            else:
                record.update(semantic_role="abstract_heading", language=language(text))
        elif record.get("role") == "keywords":
            ranges = keyword_label_ranges(text)
            lbl = text[ranges["label"][0] : ranges["label"][1]] if ranges else text
            record.update(semantic_role="keywords", language=language(lbl))
        elif record.get("module") == "abstract" and record.get("role") == "module_body":
            record.update(semantic_role="abstract_body")
        elif record.get("module") == "declaration" and record.get("role") == "module_body":
            if re.match(r"^(?:本人|作者|我们)?(?:郑重|特此|谨此)?声明[:：]$", compact):
                record["semantic_role"] = "declaration_salutation"
            elif re.search(r"(?:签名|签字|日期)[:：]|\d{4}年.*月", compact):
                record["semantic_role"] = "declaration_signature"
        elif record.get("module") == "references":
            if MODULE_MARKERS["references"].fullmatch(text):
                record["semantic_role"] = "reference_heading"
            elif record.get("role") in {"reference", "list", "body", "module_body"}:
                record["semantic_role"] = "reference_item"
        elif record.get("role") == "toc":
            if MODULE_MARKERS["toc"].fullmatch(text):
                record["semantic_role"] = "toc_heading"
        elif MODULE_MARKERS["references"].fullmatch(text):
            record["semantic_role"] = "reference_heading"
        elif MODULE_MARKERS["declaration"].fullmatch(text):
            record["semantic_role"] = "declaration_heading"
        elif MODULE_MARKERS["toc"].fullmatch(text):
            record["semantic_role"] = "toc_heading"
        elif MODULE_MARKERS["acknowledgements"].fullmatch(text):
            record["semantic_role"] = "acknowledgements_heading"
        elif MODULE_MARKERS["appendix"].fullmatch(text):
            record["semantic_role"] = "appendix_heading"
        elif MODULE_MARKERS["symbols"].fullmatch(text):
            record["semantic_role"] = "symbols_heading"

    # A thesis title (including colon subtitles or 2-line title+subtitle) immediately
    # before an abstract marker belongs to that abstract, even though the old module
    # detector still calls it cover, declaration, or keywords.
    first_abs_idx = next(
        (
            idx
            for idx, rec in enumerate(records)
            if rec.get("semantic_role") in {"abstract_heading", "abstract_inline"}
        ),
        None,
    )
    allow_cover_title = (
        not _has_separate_cover(index, index.top_ids, first_abs_idx)
        if first_abs_idx is not None
        else False
    )
    for i, record in enumerate(records):
        if record.get("semantic_role") not in {"abstract_heading", "abstract_inline"}:
            continue
        lang = record["language"]
        prev_idx = i - 1
        while prev_idx >= 0 and i - prev_idx <= 4 and not records[prev_idx]["text"].strip():
            prev_idx -= 1
        if prev_idx >= 0 and _is_abstract_title_candidate(
            records[prev_idx], allow_cover_title=allow_cover_title
        ):
            previous = records[prev_idx]
            prev_text = previous["text"].strip()
            previous.update(semantic_role="abstract_title", language=lang)
            prev2_idx = prev_idx - 1
            while (
                prev2_idx >= 0
                and prev_idx - prev2_idx <= 2
                and not records[prev2_idx]["text"].strip()
            ):
                prev2_idx -= 1
            if prev2_idx >= 0 and _is_abstract_title_candidate(
                records[prev2_idx], allow_cover_title=allow_cover_title
            ):
                prev2 = records[prev2_idx]
                if (
                    re.match(r"^(?:[—–―\-]{1,2}|[:：])", prev_text)
                    or prev2.get("module") == "keywords"
                    or prev2.get("role") == "title"
                ):
                    prev2.update(semantic_role="abstract_title", language=lang)
        for following in records[i + 1 :]:
            if following.get("semantic_role") in {
                "abstract_heading",
                "abstract_inline",
                "keywords",
            }:
                break
            if following.get("semantic_role") == "abstract_body":
                following["language"] = lang
            elif following.get("module") != "abstract":
                break

    for oid, record in index.records.items():
        if record["kind"] != "paragraph":
            continue
        role, module = record["role"], record.get("module")
        semantic = {
            "body": "body_text",
            "reference": "reference_entry",
            "caption": "caption",
            "title": "document_title",
            "table_text": "table_text",
        }.get(role, role)
        if role == "heading":
            semantic = "chapter_heading" if module == "chapters" else f"{module}_heading"
        elif role == "module_body":
            semantic = f"{module}_body"
        record.setdefault("semantic_role", semantic)
        record.setdefault("language", None)
        kind = object_kind(index.elements[oid], record.get("text", ""))
        record["object_kind"] = kind
        if kind != "text":
            record["role"] = kind
            record["semantic_role"] = kind


def object_kind(node, text):
    tags = {n.tag for n in node.iter()}
    if qn("w:drawing") in tags or qn("w:pict") in tags:
        return "mixed_content" if text.strip() else "image"
    if any(tag.endswith("}oMath") or tag.endswith("}oMathPara") for tag in tags):
        return "mixed_content" if text.strip() else "equation"
    return "text"
