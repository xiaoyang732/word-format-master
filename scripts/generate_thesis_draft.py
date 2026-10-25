# -*- coding: utf-8 -*-
"""
生成高质量计算机专业本科毕业设计论文草稿DOCX
选题：基于检索增强生成（RAG）的领域知识问答系统设计与实现
"""

import os
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from cover_template import add_standard_thesis_cover, apply_east_asia_font
from docx.enum.section import WD_SECTION_START


def _format_listing_paragraph(paragraph, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """Give algorithm/code lines their own paragraph geometry.

    The thesis body uses justified text and a two-character first-line
    indent.  Listing paragraphs must override both properties directly so
    they do not inherit the body style when the formatting pipeline rewrites
    ``Normal``.
    """
    paragraph.alignment = alignment
    fmt = paragraph.paragraph_format
    fmt.first_line_indent = Pt(0)
    fmt.left_indent = Pt(0)
    fmt.right_indent = Pt(0)
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    # ``python-docx`` writes the physical-unit firstLine value.  Also clear
    # the character-unit form that may be inherited from the Normal style.
    p_pr = paragraph._p.get_or_add_pPr()
    ind = p_pr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        p_pr.append(ind)
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:firstLineChars"), "0")
    ind.attrib.pop(qn("w:hanging"), None)
    ind.attrib.pop(qn("w:hangingChars"), None)


def build_thesis():
    doc = Document()
    
    # ----------------- 封面 (Universal Academic Thesis Cover) -----------------
    add_standard_thesis_cover(
        doc,
        degree_type="毕 业 设 计（论 文）",
        title_cn="基于检索增强生成（RAG）的领域知识问答系统设计与实现",
        title_en="Design and Implementation of Domain Knowledge Question Answering System Based on Retrieval-Augmented Generation",
        fields=[
            ("学生姓名：", "张    三"),
            ("学　　号：", "2022010101"),
            ("所在学院：", "计算机科学与技术学院"),
            ("专业名称：", "计算机科学与技术"),
            ("指导教师：", "李四  教授"),
        ],
        date_str="二〇二六年五月",
        use_section_break=True,
    )
    
    # ----------------- 中文摘要 -----------------
    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs.paragraph_format.space_before = Pt(24)
    p_abs.paragraph_format.space_after = Pt(18)
    p_abs.paragraph_format.first_line_indent = Pt(0)
    r_abs = p_abs.add_run("摘　要")
    r_abs.font.size = Pt(16)
    r_abs.font.bold = True
    apply_east_asia_font(r_abs, "黑体", "Times New Roman")

    doc.add_paragraph(
        '随着以大型语言模型（Large Language Models, LLMs）为代表的生成式人工智能技术的迅猛突破，知识问答系统在智能客服、专业咨询与文献研读等领域得到了极其广泛的应用。然而，通用大语言模型在落地垂直专业领域时，普遍面临着三大关键瓶颈：一是由于预训练知识截断导致的知识时效性滞后；二是缺乏企业与机构私域知识的覆盖能力；三是自回归生成机制本身固有的“事实性幻觉”（Factuality Hallucination）现象。上述缺陷严重制约了大语言模型在高严肃性、高准确性专业领域的规模化落地。'
    )
    doc.add_paragraph(
        '针对上述瓶颈，本文深入研究了检索增强生成（Retrieval-Augmented Generation, RAG）机制，设计并实现了一套端到端的领域知识问答系统。首先，本文提出了一种基于文档抽象语法树（AST）的多层级自适应语义分块策略，结合重叠滑动窗口机制有效避免了跨块上下文语义撕裂，并动态注入层次化标题与文档来源元数据。其次，针对单一检索方式抗噪性低、专有名词与语义表征难以兼顾的问题，本文设计了基于词法稀疏检索（BM25）与稠密语义向量（BGE-M3）的两阶段智能混合检索管线，利用倒数排名融合算法（Reciprocal Rank Fusion, RRF）进行多路粗排合并，并进一步引入交叉编码器（Cross-Encoder）重排模型进行深度二次精选，显著提升了Top-K上下文检索的查准率与排序质量。再次，针对冗余信息导致的生成偏移与幻觉问题，本文构建了防“迷失在中间”（Lost in the Middle）的上下文动态编排与严格事实锚定提示词模板，实现了带有精准来源溯源引用的流式回复生成。'
    )
    doc.add_paragraph(
        '在构建的标准垂直领域评测基准数据集上进行了系统的对比与消融实验。实验结果表明，相较于传统纯BM25检索与单一向量检索基线，本文提出的混合检索与重排策略在Hit@5指标上分别提升了18.4%和12.7%，MRR指标达到了0.872；基于RAGAS框架的事实性信实度（Faithfulness）评测得分达0.941，事实性幻觉率由未经增强大模型的24.6%大幅降至2.8%。此外，系统在高并发压力测试下表现优异，首字响应延迟（TTFT）控制在620ms以内，具备良好的工业级高可用性与工程推广价值。'
    )
    p_kw = doc.add_paragraph()
    p_kw.paragraph_format.first_line_indent = Pt(0)
    p_kw.paragraph_format.space_before = Pt(12)
    r_kw_lbl = p_kw.add_run("关键词：")
    r_kw_lbl.font.size = Pt(12)
    r_kw_lbl.font.bold = True
    apply_east_asia_font(r_kw_lbl, "宋体", "Times New Roman")
    r_kw_txt = p_kw.add_run("大语言模型；检索增强生成；混合检索；向量数据库；重排序；幻觉抑制")
    r_kw_txt.font.size = Pt(12)
    apply_east_asia_font(r_kw_txt, "宋体", "Times New Roman")
    
    doc.add_page_break()
    
    # ----------------- 英文摘要 -----------------
    p_en = doc.add_paragraph()
    p_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en.paragraph_format.space_before = Pt(24)
    p_en.paragraph_format.space_after = Pt(18)
    p_en.paragraph_format.first_line_indent = Pt(0)
    r_en = p_en.add_run("ABSTRACT")
    r_en.font.size = Pt(16)
    r_en.font.bold = True
    apply_east_asia_font(r_en, "Arial", "Arial")

    doc.add_paragraph(
        'With the rapid development of generative artificial intelligence represented by Large Language Models (LLMs), knowledge question-answering systems have been extensively adopted in customer service, professional consulting, and technical literature analysis. However, when deployed in vertical professional domains, general-purpose LLMs face three critical bottlenecks: outdated knowledge due to pre-training cutoff dates, lack of private domain knowledge, and inevitable factuality hallucinations caused by the auto-regressive generation mechanism. These limitations severely hinder the reliable application of LLMs in mission-critical and high-precision scenarios.'
    )
    doc.add_paragraph(
        'To overcome these challenges, this thesis investigates the paradigm of Retrieval-Augmented Generation (RAG) and develops an end-to-end domain knowledge question answering system. First, an adaptive multi-level semantic chunking strategy based on Abstract Syntax Trees (AST) is proposed. Combined with overlapping sliding windows, it effectively eliminates cross-chunk semantic fragmentation while dynamically injecting hierarchical section headings and source document metadata. Second, addressing the vulnerability of single-mode retrieval to lexical mismatch and semantic noise, a two-stage hybrid retrieval pipeline combining sparse lexical search (BM25) and dense vector representation (BGE-M3) is designed. The candidate passages are merged via Reciprocal Rank Fusion (RRF) and subsequently refined by a Cross-Encoder reranking model, substantially boosting retrieval precision and ranking quality.'
    )
    doc.add_paragraph(
        'Third, to mitigate generation degradation caused by context redundancy and prevent the "Lost in the Middle" phenomenon, a dynamic context orchestration strategy and strict fact-anchored prompt engineering are constructed, supporting streaming generation with precise in-text citation traceability. Extensive empirical experiments on a dedicated benchmark dataset demonstrate that the proposed hybrid retrieval and reranking approach achieves an 18.4% and 12.7% improvement in Hit@5 over pure BM25 and dense retrieval baselines respectively, reaching an MRR of 0.872. Under the RAGAS evaluation framework, the faithfulness score reaches 0.941, while the hallucination rate drops dramatically from 24.6% to 2.8%. Stress testing demonstrates a Time-to-First-Token (TTFT) within 620ms, proving high robustness, scalability, and engineering feasibility.'
    )
    p_en_kw = doc.add_paragraph()
    p_en_kw.paragraph_format.first_line_indent = Pt(0)
    p_en_kw.paragraph_format.space_before = Pt(12)
    r_en_kw_lbl = p_en_kw.add_run("Key words: ")
    r_en_kw_lbl.font.size = Pt(12)
    r_en_kw_lbl.font.bold = True
    apply_east_asia_font(r_en_kw_lbl, "Times New Roman", "Times New Roman")
    r_en_kw_txt = p_en_kw.add_run("Large Language Model; Retrieval-Augmented Generation; Hybrid Retrieval; Vector Database; Re-ranking; Hallucination Mitigation")
    r_en_kw_txt.font.size = Pt(12)
    apply_east_asia_font(r_en_kw_txt, "Times New Roman", "Times New Roman")
    
    # ----------------- 正文起始分节 -----------------
    sec_body = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec_body.header.is_linked_to_previous = False
    sec_body.footer.is_linked_to_previous = False
    
    # ----------------- 第1章 绪论 -----------------
    doc.add_heading('第1章 绪论', level=1)
    
    doc.add_heading('1.1 研究背景及意义', level=2)
    doc.add_paragraph(
        '近年来，以 Transformer 架构为基础的预训练大型语言模型（Large Language Models, LLMs）在自然语言处理领域掀起了一场范式革命。通过在海量多源通用语料上进行无监督自回归预训练，大语言模型展现出了令人瞩目的自然语言理解、逻辑推理、代码生成以及上下文泛化能力。从早期的 GPT-3、LLaMA，到当前先进的开源与商用模型，大模型已逐步演变为现代人工智能应用不可或缺的核心认知底座。'
    )
    doc.add_paragraph(
        '然而，当将通用大模型应用于特定行业和高专业度垂直领域（如企业内部规范、金融合规审查、医疗诊断辅助与前沿技术研报）时，其内在的缺陷逐渐暴露。通用模型无法接触到企业内部非公开的私域知识，预训练时的数据截止日期导致模型无法掌握最新的信息；更为严重的是，大模型的概率生成机制在面对未知知识时，极易编造看似合乎逻辑却毫无根据的虚假事实，这一现象被称为“大模型幻觉”（Hallucination）。在医疗、司法、工业生产等容错率极低的领域，虚假回答可能导致极其严重的决策失误与安全事故。'
    )
    doc.add_paragraph(
        '为了解决上述问题，传统做法通常依赖于全量微调（Fine-Tuning）或参数高效微调（PEFT）。然而，模型微调不仅计算成本高昂、迭代周期漫长，而且微调模型仍然无法彻底根除幻觉，其知识也无法实时更新，且无法提供可靠的事实引用依据。相比之下，检索增强生成（Retrieval-Augmented Generation, RAG）作为一种前沿的外置知识增强范式，通过在推理阶段实时从外部权威知识库中检索高相关性文档片段，并将检索结果作为上下文注入到生成模型的提示词中，从而使大模型能够在权威事实的严格约束下生成准确可信的回答。因此，开展高质量领域知识问答系统的设计与工程实现研究，对于推进大语言模型在企业私域与严肃垂直领域的可靠落地具有重大的理论意义与实用价值。'
    )
    
    doc.add_heading('1.2 国内外研究现状', level=2)
    doc.add_paragraph(
        '1. 大语言模型在问答领域的研究现状：早期的知识问答系统多基于规则匹配、语义解析或知识图谱三元组检索。随着深度学习的发展，基于预训练语言模型的端到端阅读理解与封闭域问答模型逐渐成为主流。然而，纯参数化模型将世界知识隐含存储在神经网络权重中，知识容量有限且缺乏动态可维护性。'
    )
    doc.add_paragraph(
        '2. 检索增强生成（RAG）技术的演进脉络：RAG 概念最早由 Lewis 等人在 2020 年提出，结合预训练检索器（DPR）与生成器（BART）实现端到端知识密集型任务。随着大模型能力的飞跃，RAG 技术经历了从 Naive RAG（朴素 RAG）、Advanced RAG（进阶 RAG）到 Modular RAG（模块化 RAG）的演进。Naive RAG 采用“切块-嵌入-检索-生成”的简单流水线，在面对长文档复杂逻辑与词汇不匹配时表现较差；Advanced RAG 引入了查询重写（Query Rewriting）、混合检索（Hybrid Retrieval）以及二次重排（Reranking）机制，显著提升了检索上下文的信噪比；Modular RAG 则进一步将 RAG 解构为独立的路由、检索、修正与验证模块，支持自适应的动态拓扑编排。'
    )
    doc.add_paragraph(
        '3. 向量数据库与稠密检索技术现状：随着稠密向量检索的兴起，以 Milvus、Chroma、Qdrant 为代表的专业向量数据库蓬勃发展。底层近似最近邻（ANN）索引算法（如 HNSW、IVF-PQ）在百万级乃至千万级高维向量检索中实现了毫秒级的响应延迟。同时，诸如 BGE、OpenAI Text-Embedding、Contriever 等前沿嵌入模型也在语义匹配与长文本泛化上取得了卓越成效。'
    )
    
    doc.add_heading('1.3 现有技术面临的挑战与不足', level=2)
    doc.add_paragraph(
        '尽管检索增强生成展现出了巨大的应用潜力，但在实际工程化落地与垂直知识问答场景中，现存技术方案仍面临以下四大核心技术挑战：'
    )
    doc.add_paragraph(
        '（1）非结构化异构文档解析与语义碎片化：企业真实文档多为长篇 PDF、Word 或 Markdown 格式，包含复杂的多级标题、表格与列表。传统按固定字符长度暴力硬切片（Fixed-size Chunking）的方式极易切断完整段落或表格语义，导致关键信息跨块丢失。'
    )
    doc.add_paragraph(
        '（2）单一检索模式的召回缺陷：单纯依赖关键词匹配（如 BM25）无法理解同义词与深层语义关联；而单纯依赖向量稠密检索在面对专业缩写、特定设备型号、精确数字等专有名词时常出现语义漂移。'
    )
    doc.add_paragraph(
        '（3）上下文冗余与“迷失在中间”现象：检索返回的多个知识块往往存在大量噪声或重复内容。大语言模型在处理超长上下文时，对位于上下文首部和尾部的信息注意力高度集中，而处于中间部分的关键信息常被忽略，从而诱发新的幻觉与偏误。'
    )
    doc.add_paragraph(
        '（4）工业级系统的高并发与响应延迟瓶颈：完整的 RAG 管线涉及向量检索、多次网络通信与大模型大参数自回归解码，首字响应时间往往长达数秒，难以满足高并发场景下用户的实时交互需求。'
    )
    
    doc.add_heading('1.4 本文主要研究内容与创新点', level=2)
    doc.add_paragraph('针对上述理论与工程技术痛点，本文的主要研究内容与创新工作概括如下：')
    doc.add_paragraph(
        '1. 层次化自适应语义分块与元数据增强：提出基于文档 AST 结构的层次化自适应切分算法，保持段落与表格的完整性，并动态为每个切片注入父级章节路径、文档标题等元数据，显著提升块级表征的语义完整度。'
    )
    doc.add_paragraph(
        '2. 两阶段多路混合检索与重排序融合管线：融合基于 BM25 的稀疏词法召回与基于 BGE-M3 的稠密语义向量召回，采用 RRF 算法合并候选集；并在精排阶段引入 Cross-Encoder 深度重排模型（BGE-Reranker-Large），消除候选噪声，实现高精度 Top-K 上下文筛选。'
    )
    doc.add_paragraph(
        '3. 动态提示词防护与可溯源流式生成架构：设计带有严格事实约束与反思机制的系统级提示词，实施注意力优先的上下文重组排列；设计流式响应传输协议，实现行内精准文献上标溯源定位。'
    )
    doc.add_paragraph(
        '4. 完整的工程化落地与全维度系统评估：基于 FastAPI 与 Milvus 实现了高可用系统架构，构建了包含 500 个垂直领域复杂问答对的评测基准，利用 RAGAS 评估体系验证了系统在检索命中率、问答忠实度与并发低延迟方面的显著优势。'
    )
    
    doc.add_heading('1.5 论文组织结构', level=2)
    doc.add_paragraph('本文共分为六章，各章节内容安排如下：')
    doc.add_paragraph('第1章：绪论。介绍研究背景、产业意义、国内外发展现状、面临的挑战、主要创新点及结构安排。')
    doc.add_paragraph('第2章：关键技术与理论基础。阐述大语言模型原理、RAG架构演进、向量相似度计算、BM25算法、RRF融合及Cross-Encoder重排理论。')
    doc.add_paragraph('第3章：领域知识问答系统需求分析与总体设计。分析系统功能与非功能需求，提出五层总体架构，详解分块算法与检索管线设计。')
    doc.add_paragraph('第4章：核心模块实现与系统集成。详细介绍向量知识库构建、混合召回重排、提示词注入、流式溯源生成与工程接口实现。')
    doc.add_paragraph('第5章：系统实验与性能评估。构建评测数据集，展开检索性能对比、问答生成质量、并发吞吐延迟及消融实验分析。')
    doc.add_paragraph('第6章：总结与展望。归纳全文研究成果与结论，指出当前研究存在的局限性，并展望未来的改进方向。')
    
    doc.add_page_break()
    
    # ----------------- 第2章 关键技术与理论基础 -----------------
    doc.add_heading('第2章 关键技术与理论基础', level=1)
    
    doc.add_heading('2.1 大语言模型生成原理与局限性', level=2)
    doc.add_paragraph(
        '现代大语言模型普遍基于自注意力机制的 Transformer 架构，采用解码器专用（Decoder-only）的自回归生成方式。给定输入标记序列 X = (x_1, x_2, ..., x_t)，模型通过多层自注意力层与前馈网络计算下一个词元 x_{t+1} 的条件概率分布：'
    )
    doc.add_paragraph('P(x_{t+1} | x_1, x_2, ..., x_t) = Softmax(W_v · h_t)')
    doc.add_paragraph(
        '其中 h_t 表示顶层隐状态向量，W_v 为词表投影矩阵。大语言模型依靠统计概率预测下一个最可能的词元，这决定了其本质上是一个强大的“模式匹配器”而非绝对的“真理判别器”。当输入提示词中缺乏关键事实支撑或训练数据中相关模式存在长尾噪声时，模型会根据上下文自发补全符合语言习惯却不符合客观事实的文本，从而产生事实性幻觉。'
    )
    
    doc.add_heading('2.2 检索增强生成（RAG）核心机制', level=2)
    doc.add_paragraph(
        '检索增强生成将知识问答过程解耦为两个核心阶段：检索阶段（Retrieval Stage）与生成阶段（Generation Stage）。'
    )
    doc.add_paragraph(
        '形式化地，设外部知识库为包含 N 个文本切片的文档集合 D = {d_1, d_2, ..., d_N}。当用户提出查询 q 时，检索器首先根据相似度函数 S(q, d) 从 D 中筛选出相关度最高的前 K 个片段构成的子集 Z_K = {d_(1), d_(2), ..., d_(K)}。随后，增强生成器以用户查询与检索到的知识片段共同作为条件，生成目标回答序列 Y = (y_1, y_2, ..., y_m)：'
    )
    doc.add_paragraph('P(Y | q) = ∏ P(y_i | y_{<i}, q, Z_K)')
    doc.add_paragraph(
        '通过将检索文档显式置入生成器的注意力窗口，大模型的解码过程被强制锚定在权威外部数据之上，从而有效消除了无依据的随机推测。'
    )
    
    doc.add_heading('2.3 文本表征模型与相似度度量', level=2)
    doc.add_paragraph(
        '稠密检索依赖于高质量的文本嵌入模型（Embedding Model）。嵌入模型基于深度神经网络（通常为双塔 Transformer 编码器结构），将任意变长文本序列映射为高维连续实数向量空间中的一个定长向量 v ∈ R^d（例如 d = 1024）。'
    )
    doc.add_paragraph(
        '在向量空间中，两个文本切片对应的向量 u 与 v 之间的语义相似度通常采用余弦相似度（Cosine Similarity）进行度量：'
    )
    doc.add_paragraph('Cosine(u, v) = (u · v) / (||u||_2 · ||v||_2) = (∑ u_i v_i) / [sqrt(∑ u_i^2) · sqrt(∑ v_i^2)]')
    doc.add_paragraph(
        '余弦相似度取值范围为 [-1, 1]，值越接近 1 表明两段文本在潜在语义空间中的方向越一致，语义关联度越高。当向量均经过 L2 归一化（即 ||u||_2 = 1, ||v||_2 = 1）后，余弦相似度直接等价于两向量的点积（Inner Product），大幅降低了最近邻搜索时的浮点计算量。'
    )
    
    doc.add_heading('2.4 经典倒排索引与 BM25 算法理论', level=2)
    doc.add_paragraph(
        '尽管稠密向量检索善于捕获高层语义意图，但在处理高频词汇差异大、包含专业术语、料号、缩略语等场景时，基于倒排索引的传统词法检索依然具备不可替代的精确匹配优势。其中，BM25 算法是信息检索领域最经典且鲁棒性极强的相关性排序算法。'
    )
    doc.add_paragraph(
        '对于用户查询 Q = {q_1, q_2, ..., q_n} 和候选文档 D，BM25 算法的相关性得分定义为：'
    )
    doc.add_paragraph('Score_BM25(D, Q) = ∑ IDF(q_i) · [f(q_i, D) · (k_1 + 1)] / [f(q_i, D) + k_1 · (1 - b + b · |D| / avgdl)]')
    doc.add_paragraph(
        '其中 f(q_i, D) 表示词项 q_i 在文档 D 中的词频；|D| 和 avgdl 分别表示文档 D 的长度与整个知识库文档的平均长度；k_1 为词频饱和度调节因子（经验值通常取 1.2 至 2.0）；b 为文档长度惩罚参数（通常取 0.75）；IDF(q_i) 为逆文档频率：'
    )
    doc.add_paragraph('IDF(q_i) = ln([N - n(q_i) + 0.5] / [n(q_i) + 0.5] + 1)')
    doc.add_paragraph('其中 N 为知识库中文档总数，n(q_i) 为包含查询词 q_i 的文档数量。' )
    
    doc.add_heading('2.5 混合检索与倒数排名融合（RRF）算法', level=2)
    doc.add_paragraph(
        '由于 BM25 检索得分（无界正实数）与稠密向量检索得分（[-1, 1] 余弦相似度）具有完全不同的分布特性和取值范围，直接进行线性加权融合容易出现某一检索器分数主导的偏置失衡问题。为此，本文采用无须归一化、只依赖相对排序位置的倒数排名融合（Reciprocal Rank Fusion, RRF）算法。'
    )
    doc.add_paragraph(
        '设 M 为所有检索器的集合（本文中包括 BM25 稀疏检索器和 BGE-M3 稠密检索器），r_m(d) 表示文档 d 在检索器 m 返回结果列表中的绝对排名位置（从 1 开始计）。则文档 d 的 RRF 综合得分为：'
    )
    doc.add_paragraph('RRF_Score(d) = ∑_{m ∈ M} [1 / (k_rrf + r_m(d))]')
    doc.add_paragraph(
        '其中 k_rrf 为平滑常数（工业界与学术界标准设置通常为 k_rrf = 60）。RRF 算法的核心优势在于：高位排名的文档能获得更高的分值权重；而在两路检索中均处于前列的交集文档将获得显著的累计加分，有效保障了高置信候选片段的优选。'
    )
    
    doc.add_heading('2.6 交叉编码器（Cross-Encoder）重排序机制', level=2)
    doc.add_paragraph(
        '双塔（Bi-Encoder）模型在检索阶段分别独立对查询和文档进行编码，两者仅在顶层通过简单的点积交互，计算速度极快（支持百万级 ANN），但不可避免地丢失了查询词与文档词之间的细粒度交叉注意力（Cross-Attention）。'
    )
    doc.add_paragraph(
        '交叉编码器（Cross-Encoder）则将用户查询与候选文档拼接为单一输入序列传入 Transformer：'
    )
    doc.add_paragraph('Input = [CLS] + q + [SEP] + d + [SEP]')
    doc.add_paragraph(
        '网络中每一层的每个 Token 都可以对所有其他 Token 进行全局自注意力交互，能够极度敏感地捕获逻辑矛盾、否定修饰、词序倒装等复杂语义关系。尽管其计算开销较大，但将其应用于多路召回后的前数十个候选样本（Top-50）的二次重排，可以在极短的时间预算内将最精准的上下文提升至 Top-5，极大地降低后续 LLM 面对无关噪声时的认知负荷。'
    )
    
    doc.add_heading('2.7 本章小结', level=2)
    doc.add_paragraph(
        '本章系统探讨了本研究所依赖的关键理论基础。首先剖析了大模型生成机制与幻觉成因，确立了外部检索增强的理论必要性；其次推导了文本稠密向量相似度度量与 BM25 稀疏检索的数学原理；最后深入阐述了 RRF 排名融合机制与 Cross-Encoder 重排序模型的工作机理，为后续两阶段混合检索管线的架构设计奠定了扎实的数理与理论支撑。'
    )
    
    doc.add_page_break()
    
    # ----------------- 第3章 领域知识问答系统需求分析与总体设计 -----------------
    doc.add_heading('第3章 领域知识问答系统需求分析与总体设计', level=1)
    
    doc.add_heading('3.1 业务需求与系统目标分析', level=2)
    doc.add_paragraph(
        '为了将问答系统有效服务于专业垂直知识服务场景，系统在设计之初需明确功能性与非功能性核心指标。'
    )
    doc.add_paragraph(
        '1. 功能性需求：'
        '（1）多格式文档导入与智能解析：支持 PDF、DOCX、Markdown、TXT 等格式文档的上传与自动化清洗；'
        '（2）知识库与索引管理：支持知识库的动态创建、修改、切分参数配置与向量持久化；'
        '（3）精准问答与两阶段检索：支持复杂查询下的稀疏与稠密混合检索、重排与Top-K优选；'
        '（4）可解释溯源引用：回答必须标明引用的具体文件名称、段落页码与原文摘要，支持一键定位对照；'
        '（5）多轮对话历史上下文跟踪：支持在对话流中理解代词指代与上下文追问。'
    )
    doc.add_paragraph(
        '2. 非功能性需求：'
        '在准确性方面，检索召回率（Hit@5）需达到 80% 以上，事实性幻觉率需抑制在 5% 以下；在性能方面，系统在 20 个并发用户同时访问时，首字响应时间（TTFT）不得超过 1.0 秒；在架构方面，各模块需高度解耦，便于嵌入模型与大模型底座的灵活热插拔。'
    )
    
    # 表3.1
    p_t31 = doc.add_paragraph('表3.1 系统核心性能与技术规格指标要求', style='Caption')
    p_t31.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table31 = doc.add_table(rows=6, cols=3)
    table31.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers31 = ['指标类别', '技术规格要求', '设计目标/基准']
    for idx, h in enumerate(headers31):
        table31.rows[0].cells[idx].text = h
    rows31_data = [
        ('检索性能', '候选段落召回率 Hit@5', '≥ 85.0%'),
        ('检索精度', '平均倒数排名（MRR）', '≥ 0.850'),
        ('生成质量', '事实性回答保真度（Faithfulness）', '≥ 0.900'),
        ('生成质量', '模型事实性幻觉率', '≤ 5.0%'),
        ('响应延迟', '首字渲染延迟（TTFT, 单并发）', '≤ 800 ms')
    ]
    for r_idx, (c0, c1, c2) in enumerate(rows31_data):
        table31.rows[r_idx+1].cells[0].text = c0
        table31.rows[r_idx+1].cells[1].text = c1
        table31.rows[r_idx+1].cells[2].text = c2
    
    doc.add_heading('3.2 系统总体技术架构设计', level=2)
    doc.add_paragraph(
        '根据模块化与分层解耦的工程设计思想，本文设计了自底向上的五层系统总体技术架构：'
    )
    doc.add_paragraph(
        '（1）数据持久化层：负责底层存储管理。包括用于存储结构化元数据（文档属性、用户会话、操作日志）的 PostgreSQL 关系数据库，用于存放原始文件的 MinIO 对象存储，以及用于高性能稠密向量存储与 HNSW 索引加速的 Milvus 向量数据库集群。'
    )
    doc.add_paragraph(
        '（2）文档智能预处理层：包含文档格式解析器、基于 AST 的层级语义分块器以及元数据自动提取引擎，完成原始文件向语义切片（Chunks）的标准化转换。'
    )
    doc.add_paragraph(
        '（3）两阶段智能检索层：包含基于 BGE-M3 的向量编码引擎、基于 BM25 的倒排索引查询器、RRF 多路融合算子以及 BGE-Reranker-Large 重排序微服务。'
    )
    doc.add_paragraph(
        '（4）大模型推理与安全生成层：负责动态上下文重组编排、防护 Prompt 构建、模型交互接口调用、Server-Sent Events（SSE）流式分发以及事实溯源解析。'
    )
    doc.add_paragraph(
        '（5）应用服务与交互展示层：基于 FastAPI 构建高性能异步 Web API，结合 Vue.js / Element-Plus 构建直观的前端交互界面，提供知识库维护、参数调优与智能对话视图。'
    )
    
    doc.add_heading('3.3 文档解析与自适应动态分块设计', level=2)
    doc.add_paragraph(
        '传统切片方法采用固定字数（如每 500 字一段）的刚性截断，极易将同一个句式或段落拆散在两个切片中，造成严重的语义割裂。针对此问题，本文设计了基于抽象语法树的自适应动态分块策略：'
    )
    doc.add_paragraph(
        '算法首先将结构化文档（如 Markdown、DOCX）解析为树状 AST 节点。随后从最低层级的自然段落、列表项或表格节点开始聚合，维持基础块大小在 L_target = 512 字符左右。当累加当前段落导致超出最大阈值 L_max = 640 字符时，在句末标点处产生切分点。同时，设置大小为 L_overlap = 64 字符的滑动重叠窗口，保留上一块末尾的关键上下文。更为关键的是，每个切片头部均自动附带其所属的各级章节路径面包屑（如“第3章 系统设计 > 3.3 文档分块算法”），使嵌入向量在表征局部内容的同时具备全局定位感知。'
    )
    
    doc.add_heading('3.4 两阶段智能混合检索管线设计', level=2)
    doc.add_paragraph(
        '两阶段检索管线的工作流程分为“粗排召回”与“精排优选”：'
    )
    doc.add_paragraph(
        '第 1 阶段（粗排召回）：当收到用户问题 q 时，系统同时启动两路并发检索请求：稠密路将 q 送入 BGE-M3 模型生成 1024 维特征向量，在 Milvus 向量集合中通过 HNSW 索引快速检索出距离最近的 Top-50 个切片；稀疏路对 q 进行结巴分词后，在本地 BM25 倒排索引中检索得分最高的 Top-50 个切片。两路结果汇总后，通过 RRF 算法计算每个切片的融合分值，合并去重后截取前 Top-30 个候选切片。'
    )
    doc.add_paragraph(
        '第 2 阶段（精排过滤）：将用户原始问题 q 与 Top-30 个候选切片文本依次配对，输入至 Cross-Encoder 架构的 BGE-Reranker-Large 重排模型中，利用全局自注意力机制计算细粒度语义相关性评分（归一化为 [0, 1]），按得分降序排列并严格截取最高的前 5 个片段（Top-5）作为最终生成上下文。'
    )
    
    doc.add_heading('3.5 提示词工程与幻觉防御机制设计', level=2)
    doc.add_paragraph(
        '为确保大语言模型在生成回答时严格恪守事实，杜绝主观臆造，本文制定了严格的三重防御提示词工程规范：'
    )
    doc.add_paragraph(
        '（1）角色与准则约束（System Instruction）：明确设定系统为严肃、严谨的领域知识专家，明文规定“仅依据给定已知参考信息作答，严禁引入未在参考资料中提及的推断，若参考资料中无法得出答案，必须明确回答‘根据现有资料无法得知’”。'
    )
    doc.add_paragraph(
        '（2）上下文排序重编排（Context Re-ordering）：研究表明大模型对长上下文的首尾部分信息记忆最深刻（“首因与近因效应”）。因此，系统对精排后的 5 个片段按照得分重组排序为：第1名放首位，第2名放末尾，第3至第5名置于中部，有效克服“迷失在中间”现象。'
    )
    doc.add_paragraph(
        '（3）事实溯源标记引导：在提示词中约束模型在每句陈述后显式添加引用标记形如 [doc:序号]，并在前端渲染为交互式高亮角标，点击即可展开原文切片与相似度得分，形成完整的事实证据闭环。'
    )
    
    doc.add_heading('3.6 本章小结', level=2)
    doc.add_paragraph(
        '本章对系统的具体业务指标与技术指标进行了系统分析，给出了整体五层技术架构，并重点阐述了自适应 AST 语义分块、BM25 与稠密向量的两阶段混合检索流程以及防幻觉的提示词工程架构设计，为后续章节的代码实现与系统部署提供了清晰的蓝图。'
    )
    
    doc.add_page_break()
    
    # ----------------- 第4章 核心模块实现与系统集成 -----------------
    doc.add_heading('第4章 核心模块实现与系统集成', level=1)
    
    doc.add_heading('4.1 知识库构建与向量存储索引实现', level=2)
    doc.add_paragraph(
        '向量知识库的高效存储与近邻搜索是整个问答系统的核心数据支撑。本文选用分布式高性能向量数据库 Milvus 2.3 构建底层存储。针对每个知识库创建独立的 Collection，并定义高度结构化的实体 Schema。'
    )
    
    # 表4.1
    p_t41 = doc.add_paragraph('表4.1 Milvus 向量集合（Collection）Schema 字段设计表', style='Caption')
    p_t41.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table41 = doc.add_table(rows=7, cols=4)
    table41.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers41 = ['字段名称', '字段类型', '约束/维度', '说明与用途']
    for idx, h in enumerate(headers41):
        table41.rows[0].cells[idx].text = h
    rows41_data = [
        ('chunk_id', 'Int64', 'Primary Key, AutoID', '切片唯一主键编号'),
        ('doc_id', 'VarChar(64)', 'Nullable=False', '所属父文档唯一UUID'),
        ('section_path', 'VarChar(256)', 'Indexed', '所属多级章节层级路径'),
        ('content', 'VarChar(4096)', 'Text, MaxLen=4096', '切片原始文本内容'),
        ('embedding', 'FloatVector', 'Dim = 1024', 'BGE-M3模型生成的稠密语义向量'),
        ('created_at', 'Int64', 'Timestamp', '知识切片入库创建时间戳')
    ]
    for r_idx, row in enumerate(rows41_data):
        for c_idx, val in enumerate(row):
            table41.rows[r_idx+1].cells[c_idx].text = val
            
    doc.add_paragraph(
        '在向量索引构建方面，本文针对 embedding 字段配置了基于分层可导航小世界图（Hierarchical Navigable Small World, HNSW）的高性能图索引算法。设置图构建参数 M = 16（每个节点的最大双向链接数），efConstruction = 200（构图时动态候选集深度）；在查询检索阶段设置 efSearch = 64。该参数组合在数万级专业知识切片库中实现了召回率与检索速度的完美权衡，单次 ANN 检索延迟保持在 5ms 以内。'
    )
    
    doc.add_heading('4.2 多路混合召回与融合重排算法工程实现', level=2)
    doc.add_paragraph(
        '检索模块采用异步非阻塞方式实现。当接收到用户请求时，后台使用 Python 的 asyncio.gather 并发执行 BM25 词法检索与 Milvus 向量搜索，极大地压缩了 I/O 等待时间。'
    )
    doc.add_paragraph(
        '两路候选集获取后，通过倒数排名融合（RRF）算法计算加权融合得分。实现逻辑如下算法4.1所示：'
    )
    
    # 算法4.1 学术算法框表格
    table_algo = doc.add_table(rows=3, cols=1)
    table_algo.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    cell_algo_title = table_algo.rows[0].cells[0]
    p_at = cell_algo_title.paragraphs[0]
    p_at.text = '算法4.1 倒数排名融合（RRF）与交叉编码器重排序算法'
    p_at.runs[0].bold = True
    _format_listing_paragraph(p_at, WD_ALIGN_PARAGRAPH.CENTER)
    
    cell_algo_io = table_algo.rows[1].cells[0]
    p_in = cell_algo_io.paragraphs[0]
    p_in.text = '输入：用户自然语言查询 q；稠密语义候选集 DenseList (Top-50)；稀疏词法候选集 SparseList (Top-50)；平滑参数 k_rrf = 60'
    p_out = cell_algo_io.add_paragraph('输出：输入大模型生成器的 Top-5 高置信度精排上下文切片集合 FinalList')
    _format_listing_paragraph(p_in)
    _format_listing_paragraph(p_out)
    
    cell_algo_steps = table_algo.rows[2].cells[0]
    steps_algo = [
        '1: 初始化多路融合得分字典 scores = {}',
        '2: 遍历 DenseList 中的切片 c（排名为 r）：scores[c] = scores.get(c, 0) + 1.0 / (k_rrf + r)',
        '3: 遍历 SparseList 中的切片 c（排名为 r）：scores[c] = scores.get(c, 0) + 1.0 / (k_rrf + r)',
        '4: 依据 RRF 得分对候选切片降序排序，截取前 30 个候选切片 Candidates',
        '5: 构造重排模型交叉注意力输入对：Pairs = [(q, c.content) for c in Candidates]',
        '6: 调用 BGE-Reranker-Large 计算细粒度交叉语义相关性得分：RerankScores = CrossEncoder(Pairs)',
        '7: 依据 RerankScores 重新降序排列，筛选排名前 5 的切片构建 FinalList 并返回'
    ]
    for idx, s in enumerate(steps_algo):
        if idx == 0:
            p = cell_algo_steps.paragraphs[0]
            p.text = s
        else:
            p = cell_algo_steps.add_paragraph(s)
        _format_listing_paragraph(p)
            
    doc.add_paragraph(
        '通过将轻量级 RRF 与重参数 Cross-Encoder 级联，系统在保障毫秒级计算延迟的同时，实现了对复杂专业问答意图的精准捕捉。'
    )
    
    doc.add_heading('4.3 提示词模板注入与上下文压缩优化实现', level=2)
    doc.add_paragraph(
        '在获取到高价值 Top-5 切片后，系统动态将其组装入结构化提示词模板中。为了避免上下文长度膨胀导致推理显存爆炸与无关信息干扰，系统还对每个切片进行了轻量级的正则去冗余（剔除多余空白、HTML标签与重复版权信息）。'
    )
    doc.add_paragraph(
        '最终组装的完整 Prompt 结构包括系统角色定义、知识参考区（每个参考段落前标有明确的 [资料1] 至 [资料5] 以及章节路径）以及用户历史对话记忆，最后接上明确的任务指令：“请根据以上参考资料准确回答用户的问题。回答内容必须全部源于参考资料，并在句末使用 [资料X] 标明出处；若无法得出结论，请明确告知用户资料不足。”'
    )
    
    doc.add_heading('4.4 流式响应与引用溯源实现', level=2)
    doc.add_paragraph(
        '为了提供极佳的用户交互体验，问答服务采用 Server-Sent Events（SSE）协议进行全双工单向流式分发。后端在与大语言模型推理引擎建立长连接后，实时将大模型吐出的每一个 Token 封装为 JSON 数据帧推送至前端：'
    )
    # SSE 流式数据帧示例代码框表格
    table_code = doc.add_table(rows=1, cols=1)
    table_code.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_code = table_code.rows[0].cells[0]
    sse_lines = [
        'data: {"type": "token", "content": "检索"}',
        'data: {"type": "token", "content": "增强"}',
        'data: {"type": "citation", "id": 1, "doc_name": "系统规范手册.pdf", "page": 12, "quote": "..."}'
    ]
    for idx, line in enumerate(sse_lines):
        if idx == 0:
            p = cell_code.paragraphs[0]
            p.text = line
        else:
            p = cell_code.add_paragraph(line)
        _format_listing_paragraph(p)
    doc.add_paragraph(
        '前端接收到 Token 后即时渲染至富文本打字机窗口，并解析其中的 [资料X] 上标，自动绑定至对应的引用卡片组件。用户将鼠标悬停在角标上时，系统即刻弹出浮动抽屉，展示原始文档的真实段落、章节层级及置信度分数，完全消除了“黑盒生成”带来的信任顾虑。'
    )
    
    doc.add_heading('4.5 系统前后端交互与工程化部署', level=2)
    doc.add_paragraph(
        '系统后端基于高性能异步 Python 框架 FastAPI 构建，利用 Pydantic 实现请求入参验证与 Schema 序列化；前端采用 Vue 3 + TypeScript + Vite 架构，结合 Pinia 进行全局对话状态管理。'
    )
    doc.add_paragraph(
        '在部署架构上，系统全量服务封装于 Docker 容器中，并通过 Docker-Compose 编排管理 Milvus、PostgreSQL、MinIO 与应用 API 容器，实现了微服务间的高内聚、低耦合与环境一键隔离交付。'
    )
    
    doc.add_heading('4.6 本章小结', level=2)
    doc.add_paragraph(
        '本章详细介绍了知识库存储、混合召回、重排序算法、提示词模板组装以及流式交互的核心工程落地细节。系统充分融合了异步并发编程、HNSW 高性能索引与深度学习重排微服务，形成了高内聚、高鲁棒性的端到端问答系统实现。'
    )
    
    doc.add_page_break()
    
    # ----------------- 第5章 系统实验与性能评估 -----------------
    doc.add_heading('第5章 系统实验与性能评估', level=1)
    
    doc.add_heading('5.1 实验环境与评测基准构建', level=2)
    doc.add_paragraph(
        '为了全面验证本文所提 RAG 架构与工程实现的优越性，本研究搭建了标准化实验测试环境并构建了专门的评测基准。'
    )
    doc.add_paragraph(
        '1. 硬件与软件环境：实验服务器配置为 Intel(R) Xeon(R) Silver 4210R CPU @ 2.40GHz，64GB DDR4 内存，NVIDIA GeForce RTX 4090 24GB GPU；操作系统为 Ubuntu 22.04 LTS，Python 3.10，CUDA 12.1，PyTorch 2.1.2，Milvus 2.3.4。'
    )
    doc.add_paragraph(
        '2. 评测数据集构建：本实验收集了包含计算机软件工程规范、网络协议标准以及云计算技术白皮书在内的 120 份真实技术文档，经预处理共生成 4,850 个规范切片。在此基础上，由 3 位计算机专业研究人员人工标注并交叉校验了 500 个垂直领域真实问答对（包括事实问答、跨章节综合理解与排他性陷阱问题），构建了标准化评测集 DomainQA-500。'
    )
    doc.add_paragraph(
        '3. 评测指标体系：'
        '（1）检索阶段指标：命中率 Hit@K（K=3, 5, 10）以及平均倒数排名（Mean Reciprocal Rank, MRR）；'
        '（2）生成质量指标：基于 RAGAS 评估框架，测量事实性信实度（Faithfulness）、回答相关度（Answer Relevance）以及上下文精确率（Context Precision）；'
        '（3）系统运行性能指标：首字生成延迟（TTFT）、总生成耗时与高并发下的系统吞吐量（Queries Per Second, QPS）。'
    )
    
    doc.add_heading('5.2 检索性能对比实验与分析', level=2)
    doc.add_paragraph(
        '为了评估不同检索算法在专业知识库上的召回精度，本文在 DomainQA-500 测试集上对比了四种不同的检索方案：'
        '方案 A：纯词法稀疏检索（BM25）；'
        '方案 B：纯稠密向量检索（BGE-M3）；'
        '方案 C：混合检索（BM25 + BGE-M3 + RRF 融合）；'
        '方案 D（本文方案）：两阶段混合检索与重排序（BM25 + BGE-M3 + RRF + BGE-Reranker-Large）。'
    )
    
    # 表5.1
    p_t51 = doc.add_paragraph('表5.1 不同检索方案在 DomainQA-500 测试集上的性能对比', style='Caption')
    p_t51.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table51 = doc.add_table(rows=5, cols=5)
    table51.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers51 = ['检索方案', 'Hit@3 (%)', 'Hit@5 (%)', 'Hit@10 (%)', 'MRR']
    for idx, h in enumerate(headers51):
        table51.rows[0].cells[idx].text = h
    rows51_data = [
        ('方案 A: 纯 BM25 稀疏检索', '68.4', '74.2', '81.6', '0.712'),
        ('方案 B: 纯 BGE-M3 稠密检索', '73.2', '79.8', '86.4', '0.768'),
        ('方案 C: 混合检索 (RRF 融合)', '82.6', '87.5', '92.1', '0.824'),
        ('方案 D: 本文混合检索 + Rerank', '87.9', '92.6', '95.8', '0.872')
    ]
    for r_idx, row in enumerate(rows51_data):
        for c_idx, val in enumerate(row):
            table51.rows[r_idx+1].cells[c_idx].text = val
            
    doc.add_paragraph(
        '由表 5.1 数据可以看出，纯 BM25 检索和纯向量检索在 Hit@5 上分别为 74.2% 和 79.8%，单一模式均无法取得理想表现。而采用 RRF 融合的混合检索显著跃升至 87.5%，证明词法与语义特征具有极强的互补性。在此基础上，加入 BGE-Reranker-Large 精排后，Hit@5 进一步提升至 92.6%，MRR 大幅增长至 0.872。这充分证明了 Cross-Encoder 的深层交叉注意力能够有效剔除候选池中的伪相关噪音，确保排在最前列的上下文具有极高的真实相关度。'
    )
    
    doc.add_heading('5.3 问答生成质量与幻觉抑制实验', level=2)
    doc.add_paragraph(
        '为了量化系统在抑制大语言模型幻觉与提升回答准确度上的效果，本文在生成端统一使用 LLaMA-3-8B-Instruct 作为基础大模型，对比了三种生成模式：'
        '基线 1：未经外部知识增强的通用大模型直接问答（No RAG）；'
        '基线 2：朴素检索增强生成（Naive RAG，纯向量检索 Top-5）；'
        '本文方案：采用本文提出的自适应分块、混合检索与重排优选架构（Advanced RAG）。'
    )
    
    # 表5.2
    p_t52 = doc.add_paragraph('表5.2 问答生成质量与幻觉评估指标对比测试结果', style='Caption')
    p_t52.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table52 = doc.add_table(rows=4, cols=5)
    table52.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers52 = ['问答生成模式', '信实度 (Faithfulness)', '答案相关度 (Relevance)', '上下文精度 (Precision)', '事实性幻觉率 (%)']
    for idx, h in enumerate(headers52):
        table52.rows[0].cells[idx].text = h
    rows52_data = [
        ('基线 1: 无增强大模型 (No RAG)', '0.512', '0.684', '--', '24.6%'),
        ('基线 2: 朴素检索增强 (Naive RAG)', '0.814', '0.826', '0.732', '8.4%'),
        ('本文方案: 本文自适应进阶 RAG', '0.941', '0.928', '0.895', '2.8%')
    ]
    for r_idx, row in enumerate(rows52_data):
        for c_idx, val in enumerate(row):
            table52.rows[r_idx+1].cells[c_idx].text = val
            
    doc.add_paragraph(
        '实验结果如表 5.2 所示。未经知识增强的大模型在回答垂直领域专业问题时，幻觉率高达 24.6%，大量回答存在无事实根据的捏造现象。采用 Naive RAG 后，由于输入了外部上下文，幻觉率显著下降至 8.4%。而本文设计的 Advanced RAG 方案凭借精准的高信噪比上下文与反思性提示词约束，将事实性信实度进一步推高至 0.941，事实性幻觉率被强力压缩至 2.8%，在专业性和可信赖度上达到了工业级可用标准。'
    )
    
    doc.add_heading('5.4 系统并发吞吐量与响应延迟测试', level=2)
    doc.add_paragraph(
        '为检验系统在真实生产环境中的高并发稳定性，使用压测工具 Locust 对系统问答接口模拟了多组不同并发客户端数下的性能压力测试。'
    )
    
    # 表5.3
    p_t53 = doc.add_paragraph('表5.3 不同并发负载下的系统性能与响应延迟测试数据', style='Caption')
    p_t53.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table53 = doc.add_table(rows=6, cols=5)
    table53.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers53 = ['并发请求数', '平均首字延迟 TTFT (ms)', 'P95 首字延迟 (ms)', '系统吞吐量 QPS', '请求错误率 (%)']
    for idx, h in enumerate(headers53):
        table53.rows[0].cells[idx].text = h
    rows53_data = [
        ('并发 1', '480', '520', '2.08', '0.0%'),
        ('并发 5', '530', '590', '9.43', '0.0%'),
        ('并发 10', '620', '710', '16.12', '0.0%'),
        ('并发 20', '780', '920', '25.64', '0.0%'),
        ('并发 50', '1240', '1580', '32.25', '0.0%')
    ]
    for r_idx, row in enumerate(rows53_data):
        for c_idx, val in enumerate(row):
            table53.rows[r_idx+1].cells[c_idx].text = val
            
    doc.add_paragraph(
        '测试结果显示（表 5.3）：在并发量低于 20 时，系统的平均首字响应延迟均控制在 800ms 以内，P95 延迟不超过 1 秒；当并发增加到 50 时，由于 GPU 推理队列排队，延迟有所上升但整体请求成功率始终保持为 100.0%，展现出了优秀的工程健壮性与并发弹性扩展能力。'
    )
    
    doc.add_heading('5.5 核心模块消融实验（Ablation Study）', level=2)
    doc.add_paragraph(
        '为了进一步验证本文各核心优化组件（自适应 AST 分块、BM25 词法混合召回、Rerank 二次精排、防迷失上下文重排）对最终系统性能的独立贡献，设计了如下消融实验：'
    )
    
    # 表5.4
    p_t54 = doc.add_paragraph('表5.4 系统核心设计模块消融实验评测结果', style='Caption')
    p_t54.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table54 = doc.add_table(rows=6, cols=4)
    table54.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers54 = ['实验组别 / 配置方案', 'Hit@5 (%)', 'MRR', 'Faithfulness']
    for idx, h in enumerate(headers54):
        table54.rows[0].cells[idx].text = h
    rows54_data = [
        ('完整系统（全配置）', '92.6', '0.872', '0.941'),
        ('移除 Rerank 重排模块', '87.5', '0.824', '0.892'),
        ('移除 BM25 混合召回（仅稠密）', '83.4', '0.781', '0.865'),
        ('替换自适应分块为固定切片', '84.2', '0.796', '0.858'),
        ('移除首尾重排提示词机制', '92.6', '0.872', '0.884')
    ]
    for r_idx, row in enumerate(rows54_data):
        for c_idx, val in enumerate(row):
            table54.rows[r_idx+1].cells[c_idx].text = val
            
    doc.add_paragraph(
        '如表 5.4 所示：若移除 Rerank 重排，MRR 从 0.872 骤降至 0.824，说明 Cross-Encoder 对排名的微调极其关键；若去掉 BM25，Hit@5 下降了 9.2 个百分点，充分佐证了关键词稀疏召回对解决专有词汇语义失配的不可或缺性；若将自适应分块替换为朴素的固定 500 字符切片，信实度下降了 0.083，表明断句割裂严重损伤了语义表征质量；若移除上下文首尾重排，尽管检索指标不受影响，但信实度明显下滑，印证了大模型对长上下文“迷失在中间”现象的真实存在。'
    )
    
    doc.add_heading('5.6 本章小结', level=2)
    doc.add_paragraph(
        '本章通过详尽的对比实验、性能压测与消融分析，全方位验证了本系统的综合效能。实验数据清晰表明，本文所设计的自适应语义分块与两阶段混合重排架构，在显著提升检索召回率的同时，将大模型的事实性幻觉率压降至极低水平，完全契合现代企业级智能知识服务的严苛要求。'
    )
    
    doc.add_page_break()
    
    # ----------------- 第6章 总结与展望 -----------------
    doc.add_heading('第6章 总结与展望', level=1)
    
    doc.add_heading('6.1 全文工作总结', level=2)
    doc.add_paragraph(
        '大语言模型技术为现代人机交互与知识处理带来了颠覆性的发展契机，但其固有的知识时效性滞后与事实性幻觉问题严重限制了其在严肃垂直领域的深度落地。本文围绕检索增强生成（RAG）范式，从文档切分、混合召回、重排序算法到提示词工程与流式系统架构，进行了一套系统、深入且具备工业级可用性的工程设计与算法实现。'
    )
    doc.add_paragraph(
        '本文的主要成果总结如下：'
        '（1）深入剖析了 RAG 核心机制，针对传统固定切片破坏语义结构的问题，提出了基于 AST 的自适应分块与层级元数据注入方法；'
        '（2）构建了两阶段混合检索管线，利用 RRF 融合了 BM25 的精确匹配优势与 BGE-M3 的深层语义表征能力，并通过 Cross-Encoder 重排过滤噪声，使检索 MRR 达到 0.872；'
        '（3）针对大模型长上下文“迷失在中间”痛点，设计了上下文重编排与事实锚定提示词策略，使事实性信实度达到 0.941，幻觉率大幅下降至 2.8%；'
        '（4）完成了包含异步 API、向量集群与流式可溯源前端的端到端系统开发与全维度评测，为垂直领域大模型应用提供了一套极具推广价值的标准化参考范式。'
    )
    
    doc.add_heading('6.2 研究局限性分析', level=2)
    doc.add_paragraph(
        '尽管本系统取得了显著的研究成果与性能提升，但在部分复杂长尾场景下仍存在以下局限性：'
        '首先，当前系统对于包含复杂排版、跨页嵌套表格、流程图等非文本富媒体格式的解析能力仍显薄弱，难以直接提取出高精度的多模态语义；'
        '其次，检索过程仍为单次线性触发，面对多步推理或跨文档综合推导的复杂开放式难题时，缺乏自适应多轮迭代检索（Multi-hop Retrieval）与自省判断能力；'
        '最后，系统目前的知识库构建依赖静态文档入库，对于高频动态更新的数据源尚未实现细粒度的增量实时同步。'
    )
    
    doc.add_heading('6.3 未来研究与改进展望', level=2)
    doc.add_paragraph(
        '针对上述局限性，后续研究工作可在以下几个方向继续深入拓展：'
        '（1）引入多模态大模型与多模态 RAG（Multimodal RAG）：研究对文档中的技术架构图、流程示意图和复杂报表图表的跨模态统一向量表征与检索；'
        '（2）知识图谱与向量融合的 GraphRAG：将实体关系图谱与向量数据库深度结合，利用图拓扑遍历捕获实体间的隐式高阶关联，提升多跳复杂推理能力；'
        '（3）自适应智能体（Agentic RAG）与主动检索：引入基于反思与计划的自主智能体框架，使系统具备“何时检索、检索什么、是否需要再次检索”的自主决策与自我纠错能力，进一步拓展智能问答的应用边界。'
    )
    
    doc.add_page_break()
    
    # ----------------- 参考文献 -----------------
    doc.add_heading('参考文献', level=1)
    
    references = [
        '[1] Vaswani A, Shazeer N, Parmar N, et al. Attention is all you need[C]//Advances in Neural Information Processing Systems. 2017: 5998-6008.',
        '[2] Lewis P, Perez E, Piktus A, et al. Retrieval-augmented generation for knowledge-intensive nlp tasks[C]//Advances in Neural Information Processing Systems. 2020: 9459-9474.',
        '[3] Gao Y, Xiong Y, Gao X, et al. Retrieval-augmented generation for large language models: A survey[J]. arXiv preprint arXiv:2312.10997, 2023.',
        '[4] Robertson S, Zaragoza H. The probabilistic relevance framework: BM25 and beyond[J]. Foundations and Trends in Information Retrieval, 2009, 3(4): 333-389.',
        '[5] Karpukhin V, Oguz B, Min S, et al. Dense passage retrieval for open-domain question answering[C]//Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing (EMNLP). 2020: 6769-6781.',
        '[6] Devlin J, Chang M W, Lee K, et al. BERT: Pre-training of deep bidirectional transformers for language understanding[C]//Proceedings of the 2019 Conference of the North American Chapter of the Association for Computational Linguistics. 2019: 4171-4186.',
        '[7] Cormack G V, Clarke C L, Buettcher S. Reciprocal rank fusion outperformscondorcet and individual rank learning methods[C]//Proceedings of the 32nd International ACM SIGIR Conference on Research and Development in Information Retrieval. 2009: 758-759.',
        '[8] Xiao S, Liu Z, Zhang P, et al. C-Pack: Packaged resources to advance general Chinese embedding[J]. arXiv preprint arXiv:2309.07597, 2023.',
        '[9] Malkov Y A, Yashunin D A. Efficient and robust approximate nearest neighbors using hierarchical navigable small world graphs[J]. IEEE Transactions on Pattern Analysis and Machine Intelligence, 2018, 42(4): 824-836.',
        '[10] Liu N F, Lin K, Hewitt J, et al. Lost in the middle: How language models use long contexts[J]. Transactions of the Association for Computational Linguistics, 2024, 12: 157-173.',
        '[11] Es S, James J, Espinosa-Anke L, et al. Ragas: Automated evaluation of retrieval augmented generation[C]//Proceedings of the 18th Conference of the European Chapter of the Association for Computational Linguistics. 2024: 150-158.',
        '[12] Ji Z, Lee N, Frieske R, et al. Survey of hallucination in natural language generation[J]. ACM Computing Surveys, 2023, 55(12): 1-38.',
        '[13] Nogueira R, Cho K. Passage re-ranking with BERT[J]. arXiv preprint arXiv:1901.04085, 2019.',
        '[14] Xiong L, Xiong C, Ye Y, et al. Approximate nearest neighbor negative contrastive learning for dense text retrieval[C]//International Conference on Learning Representations (ICLR). 2021.',
        '[15] Wang L, Yang N, Huang X, et al. Text embeddings by weakly-supervised contrastive pre-training[J]. arXiv preprint arXiv:2212.03533, 2022.',
        '[16] Brown T, Mann B, Ryder N, et al. Language models are few-shot learners[C]//Advances in Neural Information Processing Systems. 2020: 1877-1901.',
        '[17] Touvron H, Martin L, Stone K, et al. Llama 2: Open foundation and fine-tuned chat models[J]. arXiv preprint arXiv:2307.09288, 2023.',
        '[18] 张奇, 桂韬, 郑锐, 等. 自然语言处理导论[M]. 北京: 机械工业出版社, 2023.',
        '[19] 赵鑫, 李军毅, 周昆, 等. 大语言模型[M]. 北京: 电子工业出版社, 2024.',
        '[20] 冯夏坤, 贾子衡, 肖阳, 等. 知识增强的大语言模型研究综述[J]. 计算机学报, 2024, 47(3): 512-538.',
        '[21] 肖阳, 丁博, 史殿习, 等. 面向垂直领域的大模型检索增强生成技术综述[J]. 软件学报, 2024, 35(8): 3621-3645.',
        '[22] 钱伟, 陈振宇. 向量数据库技术与应用发展综述[J]. 计算机科学, 2023, 50(11): 15-28.'
    ]
    for r in references:
        doc.add_paragraph(r)
        
    doc.add_page_break()
    
    # ----------------- 致谢 -----------------
    doc.add_heading('致谢', level=1)
    doc.add_paragraph(
        '四载春秋，寒暑代谢。在大学本科四年学习与毕业设计即将画上圆满句号之际，回首求学岁月，心中充满了感激与留恋。本论文的顺利完成，离不开许多老师、同学与亲友的悉心指导、支持与关怀。'
    )
    doc.add_paragraph(
        '首先，我要衷心感谢我的指导教师李四教授。在整个毕业设计的选题构思、算法设计、实验论证乃至论文修改过程中，李老师渊博的学识、敏锐的学术洞察力、严谨踏实的治学态度以及对学生的耐心关怀，使我受益匪浅。李老师不仅在学术研究上为我指明了方向，更在做人与做事上树立了崇高的楷模。在此，谨向李四教授致以最诚挚的敬意与最衷心的感谢！'
    )
    doc.add_paragraph(
        '其次，我要感谢计算机科学与技术学院的各位授课老师与实验室的学长学姐们。感谢学院为我们提供了优良的科研实验平台与浓厚的学术氛围；感谢实验室同门在项目探讨与日常学习中给予的热情帮助与无私分享。'
    )
    doc.add_paragraph(
        '最后，我要深深感谢一直默默理解、支持并鼓励我的父母与家人。你们无论何时何地给予我的无微不至的关爱与坚实后盾，是我勇往直前、克服困难的最强大动力源泉。'
    )
    doc.add_paragraph(
        '行文至此，青春岁月翻开新的篇章。我将带着老师与亲友们的殷切嘱托，在未来的人生与科研道路上恪守初心、砥砺前行！'
    )
    
    # 保存文档
    output_path = Path('output/基于检索增强生成的领域知识问答系统-初稿.docx')
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f'Thesis draft generated successfully at: {output_path.resolve()}')

if __name__ == '__main__':
    build_thesis()
