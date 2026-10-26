"""Conservative clause parser. Every unconsumed phrase remains reviewable."""

import re

from parse_requirements import CN_SIZE_PT, FONT_NAMES, measure_to_mm

MEASURE = r"(\d+(?:\.\d+)?)\s*(厘米|毫米|英寸|磅|cm|mm|inches|inch|in|pt)"
ROLES = [
    (r"([一二三四五六七八九1-9])级标题|heading\s*([1-9])", "heading"),
    (r"正文|body(?:\s+text)?", "body"),
    (r"摘要正文|abstract(?:\s+body)?", "abstract"),
    (r"参考文献|references|bibliography", "reference"),
    (r"关键词|keywords", "keywords"),
    (r"图题|图注|figure\s+caption", "figure"),
    (r"表题|表注|table\s+caption", "table"),
]


def parse_blocks(blocks):
    candidates, unresolved = [], []
    context = None
    scope = None
    for block in blocks:
        if block.get("scope") != scope:
            context = None
            scope = block.get("scope")
        for clause in re.split(r"[\n。；;]+", block["text"]):
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
            for pattern, role in ROLES:
                for match in re.finditer(pattern, clause, re.I):
                    selected = {"type": "paragraph", "role": role, "all": True}
                    if role == "heading":
                        raw = match.group(1) or match.group(2)
                        selected["level"] = (
                            int(raw) if raw.isdigit() else "一二三四五六七八九".index(raw) + 1
                        )
                    elif role == "abstract":
                        selected.update(role="module_body", module="abstract")
                    elif role in {"figure", "table"}:
                        selected.update(role="caption", contains="图" if role == "figure" else "表")
                    found_roles.append(selected)
                    consume(match)
            # Avoid body inside 摘要正文 being treated as two independent scopes.
            if "摘要正文" in clause:
                found_roles = [s for s in found_roles if s.get("module") == "abstract"]
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
                ("页眉距|header distance", "header_distance"),
                ("页脚距|footer distance", "footer_distance"),
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
                        r"(?:行距\s*)?(?:固定值|exact(?:ly)?)\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)",
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
                    ("font.bold.set", r"(?:不加粗|非粗体|not bold|加粗|粗体|bold)"),
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
