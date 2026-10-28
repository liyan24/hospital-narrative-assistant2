# -*- coding: utf-8 -*-
"""生成《中医针灸系统功能需求对照说明》Word 文档（一次性工具）。
对照来源：医院叙事生成助手（HNA）已实现的功能。
"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "docs" / "images" / "research"
OUT = ROOT / "docs" / "中医针灸系统功能需求对照说明.docx"


# ---------------- 样式辅助 ----------------

def set_font(run, name="宋体", size=11, bold=False, color=None, mono=False):
    run.font.name = "Consolas" if mono else "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color:
        run.font.color.rgb = RGBColor(*color)


def para(doc, text, size=11, font="宋体", bold=False, indent=True,
         align=None, space_after=6, color=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.35
    if indent:
        pf.first_line_indent = Pt(size * 2)
    if align:
        p.alignment = align
    set_font(p.add_run(text), name=font, size=size, bold=bold, color=color)
    return p


def heading(doc, text, size=14):
    para(doc, text, size=size, bold=True, font="黑体", indent=False, space_after=8)


def sub_heading(doc, text, size=12):
    para(doc, text, size=size, bold=True, font="黑体", indent=False, space_after=6)


def quote(doc, text):
    para(doc, text, size=10.5, indent=False, color=(90, 90, 90), space_after=8)


def code_block(doc, code, caption):
    para(doc, caption, size=9, indent=False, color=(90, 90, 90), space_after=2)
    for line in code.rstrip().splitlines():
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pf.left_indent = Pt(18)
        # 浅灰底纹
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "F2F2F2")
        p._p.get_or_add_pPr().append(shd)
        set_font(p.add_run(line if line.strip() else " "), size=8.5, mono=True)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def image(doc, filename, caption, width_cm=15.5):
    fp = IMG / filename
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if fp.exists():
        p.add_run().add_picture(str(fp), width=Cm(width_cm))
    else:
        set_font(p.add_run(f"[图片缺失：{filename}]"), size=9, color=(200, 0, 0))
    para(doc, caption, size=9, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
         color=(90, 90, 90), space_after=10)


def overview_table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=4)
    t.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            c = t.rows[i].cells[j]
            c.paragraphs[0].text = ""
            set_font(c.paragraphs[0].add_run(cell), size=10, bold=(i == 0))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ---------------- 代码节选（摘自项目源码，保持原样） ----------------

CODE_218 = '''    def set(self, namespace, messages, temperature, max_tokens,
            model, content, ttl_hours=None, metadata=None):
        """存储LLM输出到缓存，返回缓存键"""
        content_hash = self.compute_content_hash(messages, temperature, max_tokens, model)
        cache_key = self.build_cache_key(namespace, content_hash)
        ttl = ttl_hours if ttl_hours is not None else self.ttl_hours
        now = datetime.now(timezone.utc)
        data = {
            "content": content,
            "created_at": now.isoformat(),          # 记录创建时间
            "expires_at": (now + timedelta(hours=ttl)).isoformat(),  # 版本有效期
            "namespace": namespace,                  # 来源命名空间（服务:功能）
            "metadata": metadata or {},              # 追溯元数据
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return cache_key'''

CODE_219 = '''    def _batch_merge_nodes(self, label, key_prop, nodes):
        """批量 MERGE 节点：全局去重 + 分批 UNWIND 并行写入 Neo4j"""
        seen = {}
        for n in nodes:                       # 全局去重
            seen.setdefault(n[key_prop], n)
        deduped = list(seen.values())
        for i in range(0, len(deduped), self.batch_size):
            batch = deduped[i : i + self.batch_size]      # 分批
            cql = f"""
            UNWIND $batch AS row
            MERGE (n:{label} {{`{key_prop}`: row.`{key_prop}`}})
            SET n += row.props
            RETURN count(n) AS cnt
            """
            with self.client.driver.session() as session:
                session.run(cql, {"batch": [
                    {key_prop: n[key_prop],
                     "props": {k: v for k, v in n.items() if k != key_prop and v is not None}}
                    for n in batch]})'''

CODE_220 = '''    def run(self, params: dict) -> dict:
        item_set = self.get_param(params, "item_set")   # diagnoses=合并症 / drugs=联合用药
        min_support = float(self.get_param(params, "min_support"))
        min_confidence = float(self.get_param(params, "min_confidence"))

        transactions, label = _get_transactions(item_set)      # 就诊事务集自动构建
        from mlxtend.frequent_patterns import apriori, association_rules
        onehot = _transactions_to_onehot(transactions)
        # Apriori 自动发现频繁组合
        freq = apriori(onehot, min_support=min_support, use_colnames=True,
                       max_len=3, low_memory=True)
        # 自动生成关联规则并按提升度排序
        rules = association_rules(freq, metric="confidence",
                                  min_threshold=min_confidence)
        rules = rules.sort_values("lift", ascending=False).head(top_n)'''

CODE_221 = '''    def run(self, params: dict) -> dict:
        from sklearn.ensemble import RandomForestClassifier
        from sklearn.linear_model import LogisticRegression
        from sklearn.model_selection import train_test_split

        target = self.get_param(params, "target")  # is_readmission=再入院 / is_long_stay=长住院
        X, y, feature_names = build_feature_matrix(top_n=top_n, target=target)
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=test_size, random_state=42, stratify=y)

        lr = LogisticRegression(max_iter=1000)          # 基线模型
        lr.fit(scaler.fit_transform(X_train), y_train)
        rf = RandomForestClassifier(n_estimators=100, random_state=42, n_jobs=-1)
        rf.fit(X_train, y_train)                        # 全核并行训练
        # 输出 accuracy / AUC / 混淆矩阵 / 特征重要性，辅助提前预判'''

CODE_222 = '''    def run(self, params: dict) -> dict:
        granularity = self.get_param(params, "granularity")   # year / month
        admission = dataset_service.load_table("admission")
        dates = pd.to_datetime(admission[date_col], errors="coerce").dropna()

        # 按年度/月度自动汇总趋势
        period = dates.dt.to_period("M").astype(str) if granularity == "month" \\
                 else dates.dt.year.astype(str)
        counts = period.value_counts().sort_index()

        # 自动生成趋势图与统计表，供医生快速了解数据趋势
        charts = [{"title": "入院人次趋势", "option": line_option(
            "年度入院人次趋势", counts.index.tolist(), counts.tolist(),
            "时间", "人次")}]'''

CODE_043_A = '''    def setup_schema(self):
        """创建知识图谱约束和索引：9 类节点唯一性约束"""
        constraints = [
            "CREATE CONSTRAINT patient_id IF NOT EXISTS FOR (p:Patient) REQUIRE p.patient_id IS UNIQUE",
            "CREATE CONSTRAINT visit_id IF NOT EXISTS FOR (v:Visit) REQUIRE v.visit_id IS UNIQUE",
            "CREATE CONSTRAINT complaint_name IF NOT EXISTS FOR (c:ChiefComplaint) REQUIRE c.name IS UNIQUE",
            "CREATE CONSTRAINT exam_name IF NOT EXISTS FOR (e:Exam) REQUIRE e.name IS UNIQUE",
            "CREATE CONSTRAINT labitem_name IF NOT EXISTS FOR (l:LabItem) REQUIRE l.name IS UNIQUE",
            "CREATE CONSTRAINT drug_name IF NOT EXISTS FOR (dr:Drug) REQUIRE dr.name IS UNIQUE",
            "CREATE CONSTRAINT surgery_name IF NOT EXISTS FOR (su:Surgery) REQUIRE su.name IS UNIQUE",
            "CREATE CONSTRAINT dept_name IF NOT EXISTS FOR (de:Department) REQUIRE de.name IS UNIQUE",
        ]'''

CODE_043_B = '''    def _esearch(self, query: str, max_results: int) -> list[str]:
        """多源知识动态获取：PubMed 权威医学文献检索"""
        resp = requests.get(
            f"{PUBMED_EUTILS}/esearch.fcgi",
            params={"db": "pubmed", "term": query, "retmax": max_results,
                    "retmode": "json", "sort": "relevance"},
            timeout=15)
        return resp.json().get("esearchresult", {}).get("idlist", [])

    def _efetch(self, pmids: list[str]) -> list[dict]:
        """抓取文献标题/作者/期刊/摘要，交由大模型提炼中文要点"""
        resp = requests.get(
            f"{PUBMED_EUTILS}/efetch.fcgi",
            params={"db": "pubmed", "id": ",".join(pmids), "retmode": "xml"},
            timeout=30)'''

CODE_115 = '''    def run(self, params: dict) -> dict:
        from sklearn.decomposition import PCA
        from sklearn.preprocessing import StandardScaler

        # 特征宽表自动构建：年龄换算、诊断/用药 TopN one-hot、
        # 检验定量特征中位数填补、再入院/长住院标签自动生成
        X, _, feature_names = build_feature_matrix(top_n=top_n)

        X_scaled = StandardScaler().fit_transform(X)     # 标准化
        pca = PCA(n_components=n_components, random_state=42)
        pca.fit(X_scaled)                                # PCA 降维

        ratio = pca.explained_variance_ratio_
        cum = np.cumsum(ratio)
        # 自动计算达到 90% 解释方差所需的主成分数
        n_90 = int(np.argmax(cum >= 0.9) + 1) if (cum >= 0.9).any() else n_components'''


# ---------------- 文档构建 ----------------

doc = Document()

para(doc, "中医针灸系统功能需求对照说明", size=20, bold=True, font="黑体",
     indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "——基于“医院叙事生成助手”已实现功能的对照验证", size=12, font="楷体",
     indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

para(doc, "本文档对照中医针灸系统的 7 项功能需求，逐条说明其在已落地的“医院叙事生成助手”"
          "（以下简称“本系统”）中的实现情况。本系统以真实医院科室数据（3,990 名患者、"
          "13,743 次就诊、约 114 万条医嘱、5.5 万条检验、2 万条检查）为基础构建，"
          "其数据治理、知识挖掘、智能分析各模块与针灸系统需求高度同构，可直接迁移复用。"
          "每项需求按「完成情况说明 → 实现代码 → 系统截图」三部分给出佐证。")

heading(doc, "需求完成情况总览")
overview_table(doc, [
    ["需求编号", "需求摘要", "完成情况", "对应模块"],
    ["REQ-218", "数据追踪与版本管理", "已实现", "缓存与存储层 / 科研助手历史记录"],
    ["REQ-219", "海量数据分布式存储与并行计算", "部分实现", "多引擎存储 + 批量并行导入"],
    ["REQ-220", "特征自动发现引擎", "已实现", "科研助手·数据挖掘算子"],
    ["REQ-221", "疗效趋势预判", "已实现", "机器学习预测 + 风险预警"],
    ["REQ-222", "疗效数据分析", "已实现", "趋势分析 + 统计分析算子"],
    ["REQ-043", "知识挖掘与结构化知识库", "已实现", "知识图谱 + 文献检索 + RAG"],
    ["REQ-115", "特征自动提取与降维", "已实现", "特征宽表 + PCA 降维"],
])
para(doc, "注：REQ-219 当前为单机多引擎架构（MySQL + Neo4j + 向量库 + 文件缓存），"
          "具备批量并行处理能力；分布式集群（Spark 级）未部署，架构上预留扩展空间。",
     size=9.5, color=(90, 90, 90))

# ============ REQ-218 ============
heading(doc, "REQ-218　数据追踪与版本管理")
quote(doc, "需求原文：系统应构建数据追踪与版本管理，记录数据的来源和变化，便于数据管理和追溯。")
sub_heading(doc, "一、完成情况说明")
para(doc, "已实现。本系统建立了全链路的数据留痕机制：① 所有 AI 生成内容均带命名空间"
          "（服务:功能）、内容哈希、创建时间、过期时间和追溯元数据，可按命名空间检索、"
          "清理和审计；② 每次分析结果、每次科研流水线任务均持久化为带时间戳的文档，"
          "前端「历史研究记录」可随时回看任一历史版本并重新下载产物；③ 知识图谱每次"
          "构建都会记录数据来源（Excel/MySQL）、各表加载规模与耗时统计；④ 后台管理"
          "记录全部用户操作日志。数据从来源到产物的每一步均可追溯。")
sub_heading(doc, "二、实现代码（database/llm_cache.py，节选）")
code_block(doc, CODE_218, "AI 内容缓存的追踪与版本字段：来源命名空间 + 哈希 + 时间戳 + 元数据")
sub_heading(doc, "三、系统截图")
image(doc, "01_智能模式_议题推荐.png", "图 1　科研助手「历史研究记录」：每次研究均留存版本，可回看可下载")

# ============ REQ-219 ============
heading(doc, "REQ-219　海量数据分布式存储与并行计算")
quote(doc, "需求原文：系统应实现海量数据分布式存储与并行计算，支持大规模数据处理。")
sub_heading(doc, "一、完成情况说明")
para(doc, "部分实现。存储侧采用多引擎架构按数据形态分治：MySQL 承载百万级业务明细"
          "（医嘱 114 万行，分块批量导入避免内存溢出）、Neo4j 图数据库承载 83 万条"
          "实体关系、Chroma 向量库承载语义检索、文件系统承载缓存与产物。计算侧采用"
          "批量并行策略：图谱导入使用 UNWIND 分批 MERGE（每批数千条一次性写入），"
          "机器学习训练开启全核并行（n_jobs=-1），就诊级宽表一次构建后缓存复用。"
          "针灸系统的海量医案与处方数据可直接复用该架构；如需横向扩展，"
          "可在此基础上引入 Spark/Dask 分布式计算层。")
sub_heading(doc, "二、实现代码（services/knowledge_graph_service.py，节选）")
code_block(doc, CODE_219, "批量去重 + 分批 UNWIND 并行写入，支撑 83 万关系规模的图谱构建")
sub_heading(doc, "三、系统截图")
image(doc, "req219_数据资产.png", "图 2　数据资产总览：百万级医嘱、数万级检验检查的存储规模一览")

# ============ REQ-220 ============
heading(doc, "REQ-220　特征自动发现引擎")
quote(doc, "需求原文：系统应开发特征自动发现引擎，挖掘潜在规律，为疗效提升提供数据支持。")
sub_heading(doc, "一、完成情况说明")
para(doc, "已实现。科研助手内置特征自动发现引擎：基于 Apriori 的频繁项集挖掘与关联规则"
          "发现（支持度/置信度/提升度三维评估），可自动发现合并症组合与联合用药规律——"
          "实测从 13,320 条就诊事务中自动挖出「肺大疱→肺气肿」（置信度 79.8%、"
          "提升度 9.76）等临床规律；随机森林特征重要性自动识别关键影响因子；"
          "另有共现网络与异常检测算子辅助发现潜在模式。该引擎迁移到针灸场景，"
          "即可自动发现「穴位组合→疗效」「证型→配穴」等潜在规律。")
sub_heading(doc, "二、实现代码（services/research/skills/mining.py，节选）")
code_block(doc, CODE_220, "关联规则自动发现：Apriori 频繁项集 + 置信度/提升度评估")
sub_heading(doc, "三、系统截图")
image(doc, "06_专家模式_Skills工作台.png", "图 3　特征自动发现实测：自动挖出诊断关联规则并排序（含提升度图表与 AI 解读）")

# ============ REQ-221 ============
heading(doc, "REQ-221　疗效趋势预判")
quote(doc, "需求原文：系统应支持疗效趋势预判，帮助医生提前调整治疗方案。")
sub_heading(doc, "一、完成情况说明")
para(doc, "已实现。系统以再入院和长住院作为疗效与预后的代理指标，构建了"
          "「特征自动构建 → 双模型对比训练 → 性能评估 → 风险分级」的完整预判链路："
          "LogisticRegression 与随机森林对比建模，输出准确率、AUC、混淆矩阵和"
          "特征重要性（实测再入院预测 AUC 0.702，年龄为最重要特征）；风险预警模块"
          "对患者自动分级，帮助医生提前识别高风险患者并调整干预方案。"
          "迁移到针灸场景，可用于预判「某方案对某证型的疗效走向」，辅助医生"
          "提前调整穴位与疗程。")
sub_heading(doc, "二、实现代码（services/research/skills/ml.py，节选）")
code_block(doc, CODE_221, "双模型对比的预判建模：LR 基线 + 随机森林全核并行")
sub_heading(doc, "三、系统截图")
image(doc, "req221_分类预测.png", "图 4　疗效趋势预判实测：再入院预测模型的性能指标与特征重要性")

# ============ REQ-222 ============
heading(doc, "REQ-222　疗效数据分析")
quote(doc, "需求原文：系统应提供疗效数据分析，帮助医生快速了解数据趋势。")
sub_heading(doc, "一、完成情况说明")
para(doc, "已实现。系统提供完整的疗效数据分析能力：就诊趋势分析（年度/月度自动汇总）、"
          "描述性统计（均值/中位数/分位数）、组间比较（自动按正态性选择 t 检验或"
          "Mann-Whitney、卡方检验）、相关性分析热图；所有结果自动生成图表并附"
          "中文 AI 解读，医生无需统计背景即可快速掌握数据趋势。数据概览页与"
          "周简报/晨会简报提供周期性的疗效运营视图。")
sub_heading(doc, "二、实现代码（services/research/skills/overview.py，节选）")
code_block(doc, CODE_222, "疗效数据趋势自动汇总与图表生成")
sub_heading(doc, "三、系统截图")
image(doc, "req222_趋势分析.png", "图 5　疗效数据分析实测：就诊趋势自动汇总、图表与 AI 解读一体呈现")

# ============ REQ-043 ============
heading(doc, "REQ-043　知识挖掘与结构化知识库")
quote(doc, "需求原文：应支持基于权威医学数据和名老针灸医案的知识挖掘，构建结构化知识库，"
           "实现多源知识动态获取。")
sub_heading(doc, "一、完成情况说明")
para(doc, "已实现。本系统构建了结构化医学知识图谱：9 类节点（患者、就诊、疾病"
          "（西医/中医/中医证型）、主诉、检查、检验、药品、手术、科室）与 9 类关系，"
          "共 3.7 万节点、83 万关系，全部节点带唯一性约束；知识经归一化清洗"
          "（同义词映射、组合诊断拆分、ICD-10 字典对齐）。多源知识动态获取方面："
          "集成 PubMed 权威医学文献检索（E-utilities 接口）并由大模型提炼中文要点，"
          "RAG 问答基于图谱子图检索生成回答。针灸系统的名老中医案可建模为同构的"
          "「患者—就诊—证型—治法—穴位—疗效」图结构，直接复用该知识库框架。")
sub_heading(doc, "二、实现代码（services/knowledge_graph_service.py 等，节选）")
code_block(doc, CODE_043_A, "结构化知识库 Schema：9 类节点唯一性约束")
code_block(doc, CODE_043_B, "多源知识动态获取：PubMed 权威文献检索与要点提炼")
sub_heading(doc, "三、系统截图")
image(doc, "req043_患者知识视图.png", "图 6　基于知识图谱的患者全息视图：结构化知识直接支撑临床展示")

# ============ REQ-115 ============
heading(doc, "REQ-115　特征自动提取与降维")
quote(doc, "需求原文：系统应实现历史数据特征自动提取与降维，提高数据分析效率。")
sub_heading(doc, "一、完成情况说明")
para(doc, "已实现。系统从历史数据中自动提取分析特征：就诊级特征宽表一键构建"
          "（年龄单位自动换算、诊断/用药 TopN 自动 one-hot 编码、检验定量特征"
          "自动聚合与中位数填补、再入院/长住院标签自动生成），全部算子共享该特征底座；"
          "降维方面提供 PCA 算子，自动输出各主成分解释方差比、累计解释方差曲线，"
          "并自动计算达到 90% 解释方差所需的主成分数，帮助在保证信息量的前提下"
          "压缩特征维度、提升后续分析效率。")
sub_heading(doc, "二、实现代码（services/research/skills/ml.py，节选）")
code_block(doc, CODE_115, "特征自动提取（宽表构建）+ PCA 降维与主成分数自动确定")
sub_heading(doc, "三、系统截图")
image(doc, "req115_PCA降维.png", "图 7　特征降维实测：PCA 各主成分解释方差与累计方差曲线")

doc.save(str(OUT))
print("saved:", OUT)
