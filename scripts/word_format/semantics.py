"""Conservative front-matter identities shared by profiles and target selection.

These identities supplement the historical module/role fields. They never
provide default formatting or change document text.
"""

import re

from document_structure import MODULE_MARKERS

from .namespaces import qn


def language(text):
    return "zh" if re.search(r"[\u4e00-\u9fff]", text) else "en"


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
                (r"^(?:学生)?姓名[:：]", "cover_student"),
                (r"^(?:所属|二级)?学院[:：]", "cover_college"),
                (r"^(?:所学)?专业[:：]", "cover_major"),
                (r"^(?:指导教师|指导老师|导师)[:：]", "cover_advisor"),
                (r"^(?:学号|学生学号)[:：]", "cover_student_id"),
                (r"^(?:日期[:：])?[\d〇○零一二三四五六七八九十]{4}年.*月", "cover_date"),
                (r"^日期[:：]", "cover_date"),
            ):
                if re.search(pattern, compact):
                    record["semantic_role"] = identity
                    break
            else:
                if record.get("role") == "title":
                    record["semantic_role"] = "cover_title"
        if MODULE_MARKERS["abstract"].fullmatch(text):
            record.update(semantic_role="abstract_heading", language=language(text))
        elif record.get("role") == "keywords":
            record.update(semantic_role="keywords", language=language(text))
        elif record.get("module") == "abstract" and record.get("role") == "module_body":
            record.update(semantic_role="abstract_body")
        elif MODULE_MARKERS["references"].fullmatch(text):
            record["semantic_role"] = "reference_heading"

    # A thesis title immediately before an abstract marker belongs to that
    # abstract, even though the old module detector still calls it cover or
    # declaration. Do not reinterpret a preceding signature or full sentence.
    for i, record in enumerate(records):
        if record.get("semantic_role") != "abstract_heading":
            continue
        lang = record["language"]
        if i:
            previous = records[i - 1]
            text = previous["text"].strip()
            if (
                text
                and len(text) <= 200
                and previous.get("role") in {"body", "module_body", "title"}
                and not previous.get("semantic_role")
                and not re.search(r"[。；;！？!?：:]|签名|签字|年.*月", text)
                and not any(p.fullmatch(text) for p in MODULE_MARKERS.values())
            ):
                previous.update(semantic_role="abstract_title", language=lang)
        for following in records[i + 1 :]:
            if following.get("semantic_role") in {"abstract_heading", "keywords"}:
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
