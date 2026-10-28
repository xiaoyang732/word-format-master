"""Conservative clause parser. Every unconsumed phrase remains reviewable."""

import re

from parse_requirements import CN_SIZE_PT, FONT_NAMES, measure_to_mm

MEASURE = r"(\d+(?:\.\d+)?)\s*(厘米|毫米|英寸|磅|cm|mm|inches|inch|in|pt)"
ROLES = [
    (r"摘要正文|abstract\s+body", "abstract"),
    (r"(?:设计[（(]论文[）)]|论文|中文|英文)题目", "abstract_title"),
    (
        r"[“\"]?摘要[”\"]?字样|[“\"]摘要[”\"]|^\s*摘要(?=\s*[（(])|\bAbstract\b(?!\s+body)",
        "abstract_heading",
    ),
    (r"页眉(?!距)|header(?!\s+distance)", "header"),
    (r"页脚(?!距)|footer(?!\s+distance)", "footer"),
    (r"([一二三四五六七八九1-9])级标题|heading\s*([1-9])", "heading"),
    (r"正文|body(?:\s+text)?|(?<=[，,])\s*内容(?=采用|用)|^\s*内容(?=采用|用)", "body"),
    (r"章节题目|章节标题|标题(?=\s*(?:加黑|加粗))", "chapter_heading"),
    (r"参考文献|references|bibliography", "reference"),
    (r"关键词|keywords", "keywords"),
    (r"图题|图注|figure\s+caption", "figure"),
    (r"表题|表注|table\s+caption", "table"),
]


def _scopes(clause):
    """Prefer complete labels (摘要正文) over substrings (正文)."""
    found, occupied = [], []
    for pattern, role in ROLES:
        for match in re.finditer(pattern, clause, re.I):
            if any(match.start() < end and match.end() > start for start, end in occupied):
                continue
            occupied.append(match.span())
            found.append((match, role))
    return sorted(found, key=lambda pair: pair[0].start())


def _clauses(text):
    for clause in re.split(r"[\n。；;]+", text):
        # A comma introducing a new explicit scope starts a new instruction.
        # Other commas (宋体，小四，居中) keep the same instruction together.
        starts = [0]
        for match, _ in _scopes(clause):
            prefix = clause[: match.start()]
            comma = max(prefix.rfind("，"), prefix.rfind(","))
            if comma >= starts[-1] and not prefix[comma + 1 :].strip():
                starts.append(comma + 1)
        for start, end in zip(starts, starts[1:] + [len(clause)]):
            yield clause[start:end]


def parse_blocks(blocks):
    candidates, unresolved = [], []
    context = None
    scope = None
    applicability = None
    for block in blocks:
        if block.get("scope") != scope:
            context = None
            applicability = None
            scope = block.get("scope")
        # A paragraph in a prose specification is not a continuation of the
        # last formatting target. Tables already carry their label in the row.
        context = None
        if re.search(r"优秀.*(?:论文|设计).*简介.*(?:规格|要求)", block["text"]):
            applicability = block["text"]
        if applicability:
            unresolved.append(
                {
                    "anchor": block["anchor"],
                    "quote": block["text"],
                    "applicability": applicability,
                    "reason": "Separate document specification requires applicability resolution",
                }
            )
            continue
        for clause in _clauses(block["text"]):
            if not clause.strip():
                continue
            remaining = list(clause)
            target = context

            def consume(match):
                remaining[match.start() : match.end()] = " " * (match.end() - match.start())

            def emit(action, params, match, selection=None):
                candidates.append(
                    {
                        "action": action,
                        "params": params,
                        "target": selection or target,
                        "evidence": [
                            {"anchor": block["anchor"], "quote": clause.strip(), "confidence": 1.0}
                        ],
                    }
                )
                consume(match)

            found_roles = []
            for match, role in _scopes(clause):
                selected = {"type": "paragraph", "role": role, "all": True}
                if role == "heading":
                    raw = match.group(1) or match.group(2)
                    selected["level"] = (
                        int(raw) if raw.isdigit() else "一二三四五六七八九".index(raw) + 1
                    )
                elif role == "abstract":
                    selected = {"type": "paragraph", "semantic_role": "abstract_body", "all": True}
                elif role in {"abstract_title", "abstract_heading"}:
                    selected = {
                        "type": "paragraph",
                        "semantic_role": role,
                        "all": True,
                        "language": "en" if re.search(r"英文|\bAbstract\b", clause, re.I) else "zh",
                    }
                elif role == "chapter_heading":
                    selected.update(role="heading", module="chapters")
                elif role in {"figure", "table"}:
                    selected.update(role="caption", contains="图" if role == "figure" else "表")
                found_roles.append(selected)
                consume(match)
            found_roles = [dict(items) for items in {tuple(s.items()) for s in found_roles}]
            if len(found_roles) > 1:
                unresolved.append(
                    {
                        "anchor": block["anchor"],
                        "quote": clause,
                        "reason": "Multiple formatting scopes in one clause",
                    }
                )
                continue
            if found_roles:
                target = context = found_roles[0]

            # A bibliography example or other narrative is evidence to review,
            # never an implicit instruction to set a font named after a city.
            if re.match(r"\s*[［\[]\d+[］\]]", clause):
                target = context = None

            page = {"type": "section", "all": True}
            for match in re.finditer(r"\bA4\b|\bLetter\b", clause, re.I):
                width, height = (210, 297) if match.group().upper() == "A4" else (215.9, 279.4)
                emit("page.width.set", {"value": width}, match, page)
                emit("page.height.set", {"value": height}, match, page)
            for match in re.finditer(r"横向|纵向|landscape|portrait", clause, re.I):
                emit(
                    "page.orientation.set",
                    {
                        "value": "landscape"
                        if match.group().lower() in {"横向", "landscape"}
                        else "portrait"
                    },
                    match,
                    page,
                )
            for match in re.finditer(
                rf"(?:四周|全部|所有)?(?:页边距|margins?)\s*(?:均为|为|均|:|：)?\s*{MEASURE}",
                clause,
                re.I,
            ):
                for side in ("top", "bottom", "left", "right"):
                    emit(
                        f"page.margin_{side}.set",
                        {"value": measure_to_mm(float(match[1]), match[2])},
                        match,
                        page,
                    )
            for label, action in [
                ("上|top", "margin_top"),
                ("下|bottom", "margin_bottom"),
                ("左|left", "margin_left"),
                ("右|right", "margin_right"),
                ("装订线|gutter", "gutter"),
                ("页眉距?|header distance", "header_distance"),
                ("页脚距?|footer distance", "footer_distance"),
            ]:
                for match in re.finditer(
                    rf"(?:{label})(?:页?边距|\s+margin)?\s*(?:为|:|：)?\s*{MEASURE}", clause, re.I
                ):
                    emit(
                        f"page.{action}.set",
                        {"value": measure_to_mm(float(match[1]), match[2])},
                        match,
                        page,
                    )

            # Scoped character and paragraph rules. No default body scope.
            if target:
                for font in sorted(FONT_NAMES, key=len, reverse=True):
                    for match in re.finditer(re.escape(font), "".join(remaining), re.I):
                        emit(
                            "font.east_asia.set"
                            if re.search(r"[\u4e00-\u9fff]", font)
                            else "font.latin.set",
                            {"value": font},
                            match,
                        )
                        if (
                            target.get("role") == "body"
                            and "一律" in clause
                            and re.search(r"[\u4e00-\u9fff]", font)
                        ):
                            emit(
                                "font.east_asia.set",
                                {"value": font},
                                match,
                                {
                                    "type": "paragraph",
                                    "role": "heading",
                                    "module": "chapters",
                                    "all": True,
                                },
                            )
                for label, size in sorted(
                    CN_SIZE_PT.items(), key=lambda pair: len(pair[0]), reverse=True
                ):
                    for match in re.finditer(re.escape(label), "".join(remaining)):
                        emit("font.size.set", {"value": size}, match)
                patterns = [
                    (
                        r"(?:字号|font\s*size)\s*(?:为|:|：)?\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)",
                        "font.size.set",
                        lambda m: {"value": float(m[1])},
                    ),
                    (
                        r"(?:行距\s*)?(\d+(?:\.\d+)?)\s*倍(?:行距)?",
                        "paragraph.line_spacing.set",
                        lambda m: {"kind": "multiple", "value": float(m[1])},
                    ),
                    (
                        r"(?:行距\s*)?(?:固定值?|exact(?:ly)?)\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)",
                        "paragraph.line_spacing.set",
                        lambda m: {"kind": "exact", "value": float(m[1])},
                    ),
                    (
                        r"单倍行距|single\s+spacing",
                        "paragraph.line_spacing.set",
                        lambda m: {"kind": "multiple", "value": 1},
                    ),
                    (
                        r"双倍行距|double\s+spacing",
                        "paragraph.line_spacing.set",
                        lambda m: {"kind": "multiple", "value": 2},
                    ),
                ]
                for label, name in [
                    ("首行缩进|first[- ]line indent", "first_line_indent"),
                    ("悬挂缩进|hanging indent", "hanging_indent"),
                    ("左缩进|left indent", "left_indent"),
                    ("右缩进|right indent", "right_indent"),
                ]:
                    patterns.extend(
                        [
                            (
                                rf"(?:{label})\s*(?:为|:|：)?\s*(\d+(?:\.\d+)?)\s*(?:个)?(?:字符|chars?)",
                                f"paragraph.{name}.set",
                                lambda m: {"value": float(m[1]), "unit": "chars"},
                            ),
                            (
                                rf"(?:{label})\s*(?:为|:|：)?\s*{MEASURE}",
                                f"paragraph.{name}.set",
                                lambda m: {"value": measure_to_mm(float(m[1]), m[2]), "unit": "mm"},
                            ),
                        ]
                    )
                for label, name in [("段前|space before", "before"), ("段后|space after", "after")]:
                    patterns.append(
                        (
                            rf"(?:{label})\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)",
                            f"paragraph.spacing_{name}.set",
                            lambda m: {"value": float(m[1])},
                        )
                    )
                for pattern, action, params in patterns:
                    for match in re.finditer(pattern, "".join(remaining), re.I):
                        emit(action, params(match), match)
                for match in re.finditer(
                    r"(?:^|(?<=[，,]))\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)(?=\s*(?:[，,]|$))",
                    "".join(remaining),
                    re.I,
                ):
                    emit("font.size.set", {"value": float(match[1])}, match)
                for value, pattern in [
                    ("justify", r"两端对齐|justified|justify"),
                    ("center", r"居中|cent(?:er|re)(?:d)?"),
                    ("left", r"左对齐|left[- ]aligned?"),
                    ("right", r"右对齐|right[- ]aligned?"),
                ]:
                    for match in re.finditer(pattern, "".join(remaining), re.I):
                        emit("paragraph.alignment.set", {"value": value}, match)
                for action, pattern in [
                    ("font.bold.set", r"(?:不加粗|不加黑|非粗体|not bold|加粗|加黑|粗体|bold)"),
                    ("font.italic.set", r"(?:不倾斜|非斜体|not italic|斜体|italic)"),
                ]:
                    for match in re.finditer(pattern, "".join(remaining), re.I):
                        emit(
                            action,
                            {"value": not bool(re.match(r"不|非|not", match.group(), re.I))},
                            match,
                        )
            residual = "".join(remaining)
            residual = re.sub(
                r"(?:使用|采用|设置为|设置|字体|字号|纸张|中文|西文|英文|为|要求|格式|font|size|paper)|[\s，,：:、()（）.]+",
                "",
                residual,
                flags=re.I,
            )
            if residual:
                unresolved.append(
                    {
                        "anchor": block["anchor"],
                        "quote": clause.strip(),
                        "remaining": residual,
                        "reason": "Unparsed wording; AI must interpret the complete clause",
                    }
                )
            # Negation/conditions not belonging to supported boolean syntax
            # invalidate all candidates in this clause, even if a positive token matched.
            if re.search(
                r"不得|禁止|不要|不使用|不采用|除外|除.*外|至少|不超过|仅限|except|unless|at least|at most|\bnot\b(?!\s+(?:bold|italic))",
                clause,
                re.I,
            ):
                candidates = [
                    c
                    for c in candidates
                    if c["evidence"][0]
                    != {"anchor": block["anchor"], "quote": clause.strip(), "confidence": 1.0}
                ]
                unresolved.append(
                    {
                        "anchor": block["anchor"],
                        "quote": clause.strip(),
                        "reason": "Conditional or negative instruction requires semantic resolution",
                    }
                )
    return candidates, unresolved
