# -*- coding: utf-8 -*-
"""生成《ZJ-TC 系列用例接口 curl 测试手册》Word 文档（一次性工具）。"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "ZJ-TC系列用例接口curl测试手册.docx"


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


def sub_heading(doc, text, size=11.5):
    para(doc, text, size=size, bold=True, font="黑体", indent=False, space_after=4)


def code_block(doc, code, caption=None):
    if caption:
        para(doc, caption, size=9, indent=False, color=(90, 90, 90), space_after=2)
    for line in code.rstrip().splitlines():
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_after = Pt(0)
        pf.line_spacing = 1.05
        pf.left_indent = Pt(14)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "F2F2F2")
        p._p.get_or_add_pPr().append(shd)
        set_font(p.add_run(line if line.strip() else " "), size=9, mono=True)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


doc = Document()

para(doc, "ZJ-TC 系列用例接口 curl 测试手册", size=20, bold=True, font="黑体",
     indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "——ZJ-TC-115-001 ~ ZJ-TC-222-001 功能的后台接口直接验证方法", size=12,
     font="楷体", indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

para(doc, "本文档列出 ZJ-TC-115-001 至 ZJ-TC-222-001 共 12 个测试用例所对应功能的"
          "后台接口 curl 命令。使用 curl 直接调用后台接口，可以不依赖前端页面，"
          "快速获取各功能的原始返回数据（JSON），用于接口级验收测试。")

# ================= 通用说明 =================
heading(doc, "一、通用说明")
para(doc, "1. 接口地址：后端服务地址为 http://localhost:8005，所有接口无需登录态即可调用"
          "（/api/research、/api/narrative、/api/kg 未强制鉴权）；后台管理类接口需要"
          "先登录获取 token。", indent=False)
para(doc, "2. 超时设置：带 AI 解读的算子接口耗时约 30~120 秒，务必加 --max-time 180。", indent=False)
para(doc, "3. 中文参数：请求体含中文时，建议将 JSON 写入文件后用 --data-binary @body.json"
          " 发送，避免 Windows 终端编码问题。", indent=False)
para(doc, "4. 返回结构：算子类接口统一返回 status / result_id / summary（文字结论）/ "
          "tables（表格）/ charts（ECharts 图表配置）/ facts（关键数字）/ "
          "interpretation（AI 中文解读）。", indent=False)
para(doc, "5. 结果美化：可用管道接 python -m json.tool 格式化输出；保存证据加 -o result.json。",
     indent=False)

code_block(doc, 'curl localhost:8005/health', "前置检查：确认后端服务在运行（返回 {\"status\":\"ok\"}）")

# ================= 各用例 =================
heading(doc, "二、各用例 curl 命令")

CASES = [
    ("ZJ-TC-115-001　历史特征提取与降维",
     "POST /api/research/skills/dimensionality_reduction/run",
     '''curl -X POST localhost:8005/api/research/skills/dimensionality_reduction/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{}}' --max-time 180''',
     "summary 含主成分数与方差比结论；facts 含 pc1_ratio、cumulative_ratio、"
     "components_for_90pct（90% 方差所需主成分数）；charts 含解释方差柱状图与累计方差折线图。"),

    ("ZJ-TC-116-001　周期模式识别",
     "POST /api/research/skills/trend_analysis/run（月度粒度）+ clustering/run（聚类分群）",
     '''curl -X POST localhost:8005/api/research/skills/trend_analysis/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{"granularity":"month"}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/clustering/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{}}' --max-time 180''',
     "trend_analysis 的 charts 为月度趋势折线图，可观察周期波动；clustering 返回各群样本量、"
     "画像对比表与 PCA 散点图。"),

    ("ZJ-TC-116-002　未来趋势预测",
     "POST /api/research/skills/classification/run + regression/run",
     '''curl -X POST localhost:8005/api/research/skills/classification/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{"target":"is_long_stay"}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/regression/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{}}' --max-time 180''',
     "分类返回 accuracy / AUC / 混淆矩阵；回归返回 MAE / R²；两者 facts 均含特征重要性数据。"),

    ("ZJ-TC-117-001　预测置信度评估",
     "POST /api/research/skills/classification/run + group_comparison/run",
     '''curl -X POST localhost:8005/api/research/skills/classification/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/group_comparison/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{}}' --max-time 180''',
     "classification 的 tables 含双模型 accuracy/AUC 与混淆矩阵；group_comparison 的 "
     "tables 含各变量检验方法（t 检验/Mann-Whitney/卡方）与 p 值；interpretation 含局限性提示。"),

    ("ZJ-TC-118-001　趋势分析可视化",
     "POST trend_analysis/run + correlation/run + cooccurrence_network/run",
     '''curl -X POST localhost:8005/api/research/skills/trend_analysis/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/correlation/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/cooccurrence_network/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180''',
     "三个接口返回的 charts 字段分别为趋势折线图、相关性热图、共现网络图的完整 "
     "ECharts 配置 JSON，即可视化数据的直接证据。"),

    ("ZJ-TC-218-001　数据来源追踪",
     "GET /api/research/data-assets",
     '''curl localhost:8005/api/research/data-assets --max-time 60''',
     "assets.tables 列出各表名称、行数、列数、关键字段与覆盖说明（数据来源与规模的直接证据）；"
     "assets.graph 含图谱节点/关系统计；text_data 含文本数据规模。"),

    ("ZJ-TC-218-002　数据版本管理",
     "GET /api/research/auto/history → GET /api/research/auto/{job_id} → 下载产物",
     '''curl localhost:8005/api/research/auto/history --max-time 60

curl localhost:8005/api/research/auto/autojob_xxxxxxxx --max-time 60

curl -OJ localhost:8005/api/research/paper/download/科研论文_xxx.docx --max-time 60''',
     "history 返回历次研究记录（创建时间、议题、论文标题、状态、下载地址，同议题只保留最新）；"
     "按 job_id 可取回任一版本的完整内容与产物；docx 可正常下载。"),

    ("ZJ-TC-219-001　海量数据分布存储",
     "GET /api/research/data-assets + /api/data/tables + /api/kg/stats",
     '''curl localhost:8005/api/research/data-assets --max-time 60

curl localhost:8005/api/data/tables --max-time 60

curl localhost:8005/api/kg/stats --max-time 60''',
     "data-assets 可见百万级医嘱等存储规模；tables 列出 MySQL 全部表；kg/stats 返回图谱"
     "各类节点与关系计数（需 Neo4j 服务已启动）。"),

    ("ZJ-TC-219-002　并行计算",
     "POST /api/kg/build（图谱批量构建）+ 分类算子计时",
     '''curl -X POST localhost:8005/api/kg/build \\
  -H "Content-Type: application/json" \\
  -d '{"clear":false}' --max-time 1800

time curl -s -X POST localhost:8005/api/research/skills/classification/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 300''',
     "kg/build 执行分批并行写入，返回构建统计与耗时（全量约 5 分钟）；分类算子在秒级返回，"
     "体现并行计算能力。"),

    ("ZJ-TC-220-001　特征自动发现",
     "POST /api/research/skills/association_rules/run + feature_importance/run",
     '''curl -X POST localhost:8005/api/research/skills/association_rules/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/feature_importance/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180''',
     "association_rules 的 tables 含规则前项/后项/支持度/置信度/提升度，charts 含提升度图；"
     "feature_importance 返回 Top20 特征及重要性排序。"),

    ("ZJ-TC-221-001　疗效趋势预判",
     "POST classification/run + GET /api/narrative/risk-prediction",
     '''curl -X POST localhost:8005/api/research/skills/classification/run \\
  -H "Content-Type: application/json" \\
  -d '{"params":{"target":"is_readmission"}}' --max-time 180

curl "localhost:8005/api/narrative/risk-prediction?patient_id=557c91116af011f19de2d488664b66d4" \\
  --max-time 180''',
     "classification 输出预测模型性能与影响因素；risk-prediction 返回指定患者风险评分、"
     "风险等级与重点因素，辅助提前干预。"),

    ("ZJ-TC-222-001　疗效数据趋势分析",
     "POST trend_analysis/run + descriptive_stats/run + group_comparison/run",
     '''curl -X POST localhost:8005/api/research/skills/trend_analysis/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/descriptive_stats/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180

curl -X POST localhost:8005/api/research/skills/group_comparison/run \\
  -H "Content-Type: application/json" -d '{"params":{}}' --max-time 180''',
     "趋势图与统计表齐全；描述性统计含均值/中位数/分位数；组间比较含检验方法与 p 值；"
     "均附中文 AI 解读。"),
]

for title, api, cmd, expect in CASES:
    sub_heading(doc, title)
    para(doc, f"对应接口：{api}", size=10, indent=False, color=(60, 60, 60), space_after=2)
    code_block(doc, cmd)
    para(doc, f"预期返回要点：{expect}", size=10, indent=False, space_after=12)

# ================= 常见问题 =================
heading(doc, "三、常见问题")
para(doc, "1. 返回 404：后端为旧进程未加载最新模块，重启后端服务。", indent=False)
para(doc, "2. 返回 {\"detail\":\"Not Found\"} 且根路径显示 ZRAdmin：请求打到了 8000 端口的"
          "其他项目，本项目后端在 8005。", indent=False)
para(doc, "3. /api/kg/* 返回\"Neo4j连接失败\"：需先启动 Neo4j 图数据库服务。", indent=False)
para(doc, "4. 长时间无响应：带 AI 解读的接口 1~2 分钟返回属正常；若 interpretation 以"
          " [LLM调用失败] 开头，说明大模型调用失败（检查网络与 API Key）。", indent=False)
para(doc, "5. 登录获取 token（仅后台管理类接口需要）：", indent=False)
code_block(doc, '''curl -X POST localhost:8005/api/auth/login \\
  -H "Content-Type: application/json" \\
  -d '{"username":"admin","password":"admin123"}'
# 后续请求加请求头: -H "Authorization: Bearer <access_token>"''')

doc.save(str(OUT))
print("saved:", OUT)
