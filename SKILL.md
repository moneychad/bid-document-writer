---
name: bid-document-writer
description: "Parse tender docs and write/review bid responses."
version: 1.1.0
author: yuanman_ai
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [bid, tender, 标书, 招标, 投标, document-writing, compliance, procurement]
    related_skills: [word-doc-processing, ocr-and-documents, pdf, docx]
---

# 投标文件撰写与审核 (Bid Document Writer)

解析招标文件、撰写投标文件、审核投标响应的完整工作流。

## Changelog

### v1.1.0 (2026-09-01)

**新增：格式验证与页面布局检查**
- 新增 `setup.sh` 一键安装脚本（自动检测环境、国内镜像、依赖验证）
- 新增 Step 3.2.5：格式规范预定义（format_spec.json + 中文字号对照表）
- 新增 Step 4.1：格式诊断（docformat-gui / lxml直接检查 / python-docx）
- 新增 Step 4.2：页面布局验证（LibreOffice→PDF / 页码连续性 / 页眉页脚一致性 / 视觉对比）
- 新增中文字体检查方案：通过lxml读取w:rFonts/@w:eastAsia属性，绕过python-docx API限制
- 新增国内网络安装指南（封装在setup.sh中：PyPI清华镜像、GitHub ghproxy/gitclone镜像）
- 新增Pitfalls #16-19（中文字体陷阱、字号单位、PDF保真度、国内网络）
- Dependencies改为自动检查模式，加载skill时零开销
- 审核报告模板新增格式和页码问题示例

### v1.0.0 (初始版本)

- 五阶段工作流：解析→素材→撰写→审核→交付
- 六步读标法
- 结构化招标要求提取（10类信息）
- 逐章撰写与偏离表生成
- 合规性/格式/一致性审核清单
- 审核报告模板

---

## When to Use

- 用户提供招标文件（PDF/Word），需要解析要求并撰写投标文件
- 用户提供素材，需要根据招标要求组织成投标响应
- 需要审核投标文件的合规性和完整性
- 需要生成招标要求与投标响应的对照表

## Prerequisites

加载skill后先运行依赖检查，缺失时自动安装：

```bash
python3 -c "import docx,lxml,pdfplumber,pypdf,pymupdf" 2>/dev/null || bash ~/.hermes/skills/bid-document-writer/setup.sh
```

一行搞定，不缺依赖时零输出，不会浪费token。

### 依赖清单（仅供参考，无需每次阅读）

| 包 | 用途 |
|---|------|
| python-docx | Word读写 |
| lxml | 中文字体检查（eastAsia） |
| pdfplumber | PDF页码/页眉检查 |
| pypdf | PDF基础操作 |
| pymupdf | PDF渲染视觉对比 |
| LibreOffice | docx→PDF转换（可选，`setup.sh --full`） |
| docformat-gui | 公文格式诊断（可选，`setup.sh --full`） |

---

## Phase 1: 解析招标文件 (Parse Tender Document)

### Step 0: 六步读标法（先通读，再精读）

1. **看资格审查** — 确认企业是否满足投标资格
2. **看详细审查/评分标准** — 搞清楚怎么拿分
3. **看投标人须知前附表** — 实质性条款集中在这里（带星号/加粗 = 废标红线）
4. **看技术要求/规格书** — 技术方案的硬指标
5. **看商务条款/合同条件** — 付款、质保、违约
6. **看格式/装订/提交要求** — 形式合规

⚠️ **特别注意**：投标期间可能发布澄清/修改通知（补遗书），必须一并纳入分析！

### Step 1.1: 提取全文

**PDF文件**：
```python
import pymupdf
import pymupdf4llm

# 方法1：纯文本
doc = pymupdf.open('招标文件.pdf')
text = '\n'.join(page.get_text() for page in doc)

# 方法2：Markdown（保留结构）
md_text = pymupdf4llm.to_markdown('招标文件.pdf')
```

**Word文件**：
```python
from docx import Document
doc = Document('招标文件.docx')
text = '\n'.join(p.text for p in doc.paragraphs)

# 包含表格
for table in doc.tables:
    for row in table.rows:
        text += '\n' + '\t'.join(cell.text for cell in row.cells)
```

### Step 1.2: 提取结构化要求

将全文传给AI，要求提取以下结构化信息：

1. project_info: 项目基本信息（名称、编号、预算、投标截止时间、开标时间）
2. qualification_requirements: 资格要求列表（每条含：要求描述、是否必须、证明材料）
3. scoring_criteria: 评分标准（含：评分项、分值、评分细则）
4. technical_requirements: 技术要求列表（每条含：要求描述、章节号、是否实质性要求）
5. commercial_terms: 商务条款（付款方式、质保期、违约责任、合同主要条款）
6. format_requirements: 格式要求（装订方式、份数、封面要求、页码要求、字体要求）
7. response_checklist: 需要提供的资料清单（每项含：资料名称、是否必须、备注）
8. key_dates: 关键时间节点（投标截止、开标、答疑截止、保证金缴纳截止）
9. deviations: 允许偏离和不允许偏离的条款
10. special_notes: 特别注意事项和废标条款

### Step 1.3: 生成招标要求清单

将提取结果整理为**招标要求对照表**，作为后续撰写的基准：

```markdown
| 序号 | 章节号 | 要求类型 | 要求描述 | 是否必须 | 评分权重 | 响应状态 | 投标响应位置 |
|------|--------|---------|---------|---------|---------|---------|------------|
| 1 | 2.1 | 资格要求 | 具备ISO9001认证 | 是 | — | 【待补充】 | 第X章第X节 |
| 2 | 3.2 | 技术要求 | 系统可用性≥99.9% | 是 | 5分 | 【待撰写】 | 第X章第X节 |
```

---

## Phase 2: 收集素材 (Collect Materials)

基于招标要求清单，向用户明确需要补充的信息：

```
根据招标文件解析，以下信息需要您提供：

【必须提供】
1. 企业营业执照、资质证书扫描件
2. ISO9001等认证证书
3. 近3年类似项目业绩（合同/验收报告）
4. 项目团队人员简历和资质
5. 财务报表（近3年审计报告）

【需要确认】
6. 报价策略（总价/单价、折扣、优惠条件）
7. 技术方案的核心思路和亮点
8. 实施计划的时间安排
9. 售后服务承诺（响应时间、服务范围）

【可选提供】
10. 公司介绍和优势说明
11. 获奖/荣誉证书
12. 客户推荐信/感谢信
```

素材整理：
1. 统一收集到一个工作目录
2. 按类型分类（资质类、业绩类、技术类、商务类）
3. 提取关键数据（金额、日期、人员、指标等）
4. 标记信息完整性（哪些已齐全、哪些还缺）

---

## Phase 3: 撰写投标文件 (Write Bid Document)

### Step 3.2: 生成大纲

根据招标文件的章节结构，生成投标文件大纲。标准结构：

- 第一部分 商务文件：投标函、授权书、报价表、偏离表、资格证明
- 第二部分 技术文件：需求分析、技术方案、实施方案、质量保障、售后服务、偏离表
- 第三部分 附件：资质证书、业绩、人员简历、其他材料

### Step 3.2.5: 定义格式规范（格式预检）

**在撰写前**，先根据招标文件的格式要求，生成一份格式规范文件 `format_spec.json`：

```json
{
  "page": {
    "width_mm": 210,
    "height_mm": 297,
    "margin_top_mm": 37,
    "margin_bottom_mm": 35,
    "margin_left_mm": 28,
    "margin_right_mm": 26
  },
  "styles": {
    "title": {"font": "方正小标宋简体", "size_pt": 22, "bold": true},
    "heading1": {"font": "黑体", "size_pt": 16, "bold": true},
    "heading2": {"font": "楷体", "size_pt": 16, "bold": true},
    "heading3": {"font": "仿宋", "size_pt": 16, "bold": true},
    "body": {"font": "仿宋", "size_pt": 16, "line_spacing_pt": 28, "first_line_indent_chars": 2}
  },
  "format_check": {
    "standard": "custom",
    "forbidden_bold_in_body": true,
    "forbidden_italic_in_body": true,
    "forbidden_color_in_body": true,
    "page_numbers_required": true,
    "pii_check": true
  }
}
```

**中国公文字号对照表**（用于格式检查）：

| 字号 | 磅值(pt) | 半磅值(half-pt) | 常见用途 |
|------|---------|----------------|----------|
| 初号 | 42 | 84 | 封面标题 |
| 小初 | 36 | 72 | 封面副标题 |
| 一号 | 26 | 52 | 标题 |
| 小一 | 24 | 48 | 标题 |
| 二号 | 22 | 44 | 标题 |
| 小二 | 18 | 36 | 标题 |
| 三号 | 16 | 32 | 一级标题/正文 |
| 小三 | 15 | 30 | 二级标题 |
| 四号 | 14 | 28 | 三级标题 |
| 小四 | 12 | 24 | 正文 |
| 五号 | 10.5 | 21 | 注释 |

**关键**：招标文件说"宋体小四" = 字体"宋体" + 字号12pt。python-docx的 `font.name` 只返回拉丁字体(w:ascii)，中文字体存在 `w:rFonts/@w:eastAsia` 属性中，必须通过 lxml 直接读取XML才能验证。

### Step 3.2: 逐章撰写原则

1. **对照响应**：每一段内容都对应招标文件的某一条要求
2. **不遗漏**：用招标要求清单逐条核对
3. **不编造**：事实性信息（资质、业绩、人员）必须使用用户提供的真实数据
4. **标注待补充**：缺失的信息用`【待补充：XXX】`标注
5. **数据一致**：报价、工期、人员等关键数据在各章节保持一致

### Step 3.3: 生成偏离表

```markdown
| 序号 | 招标要求 | 投标响应 | 偏离说明 |
|------|---------|---------|----------|
| 1 | 质保期≥2年 | 质保期3年 | 正偏离 |
| 2 | 响应时间≤4小时 | 响应时间≤2小时 | 正偏离 |
```

---

## Phase 4: 审核投标文件 (Review Bid Document)

### 审核清单

**合规性审核**：
- □ 实质性要求全部响应（不满足即废标）
- □ 资格证明文件齐全
- □ 报价在预算范围内
- □ 投标保证金已缴纳
- □ 签字盖章完整
- □ 投标有效期符合要求
n- □ 无超出允许范围的偏离

**格式审核**：
- □ 封面格式符合要求
- □ 目录页码正确
- □ 字体字号符合要求
- □ 装订方式符合要求
- □ 正副本份数正确
- □ 页码连续无遗漏

**一致性审核**：
- □ 投标总价（报价表 = 商务条款 = 分项报价之和）
- □ 项目工期（技术方案 = 实施计划 = 商务条款）
- □ 人员配置（技术方案 = 人员简历表 = 业绩证明）
- □ 技术指标（技术方案 = 技术偏离表 = 参数响应表）
- □ 服务承诺（售后服务方案 = 商务条款 = 质保期）

### Step 4.1: 格式诊断（docformat-gui + lxml）

#### 方法一：docformat-gui 一键诊断

```bash
# 格式诊断模式（只报告问题，不修改文件）
cd /opt/docformat-gui
python main.py --diagnosis input.docx --config format_spec.json
```

docformat-gui 会检查：
- 页面边距（4边，cm→twips）
- 字体名称（中文字体：宋体/仿宋/黑体/楷体等）
- 字号（支持中文字号名：小四/三号等，也支持pt）
- 行距（固定值/多倍行距）
- 首行缩进
- 段前段后间距
- 对齐方式（左对齐）
- 禁止加粗/斜体/下划线/彩色（正文）
- 表格格式（边框、对齐、单元格格式）
- 分页符/分节符
- 浮动对象
- 敏感信息扫描（电话、邮箱、身份证号）

#### 方法二：lxml 直接检查中文字体

python-docx 的 `font.name` 只返回拉丁字体(w:ascii)，**中文字体存在 `w:rFonts/@w:eastAsia` 属性中**，必须通过 lxml 直接读取XML：

```python
import zipfile
from lxml import etree

def check_chinese_fonts(docx_path, expected_font='宋体', expected_size_pt=12):
    """检查docx中中文字体和字号是否符合要求"""
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    issues = []
    
    with zipfile.ZipFile(docx_path) as zf:
        with zf.open('word/document.xml') as f:
            tree = etree.parse(f)
    
    # 检查所有段落的run
    for i, para in enumerate(tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')):
        for run in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
            rpr = run.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
            if rpr is not None:
                # 检查中文字体
                rfonts = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                if rfonts is not None:
                    east_asia = rfonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                    if east_asia and east_asia != expected_font:
                        issues.append(f'段落{i+1}: 中文字体为{east_asia}，应为{expected_font}')
                
                # 检查字号（half-points，24 = 12pt = 小四）
                sz = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
                if sz is not None:
                    size_half_pt = int(sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0'))
                    expected_half_pt = int(expected_size_pt * 2)
                    if size_half_pt != expected_half_pt:
                        actual_pt = size_half_pt / 2
                        issues.append(f'段落{i+1}: 字号为{actual_pt}pt，应为{expected_size_pt}pt')
    
    return issues
```

#### 方法三：python-docx 检查页面边距和行距

```python
from docx import Document
from docx.shared import Mm, Pt, Twips
from docx.enum.text import WD_LINE_SPACING

def check_page_margins(docx_path, spec):
    """检查页面边距"""
    doc = Document(docx_path)
    issues = []
    
    for i, section in enumerate(doc.sections):
        margins = {
            'top': section.top_margin,
            'bottom': section.bottom_margin,
            'left': section.left_margin,
            'right': section.right_margin
        }
        for side, actual in margins.items():
            expected_mm = spec['page'][f'margin_{side}_mm']
            actual_mm = actual / 36000  # EMU to mm
            if abs(actual_mm - expected_mm) > 1:
                issues.append(f'节{i+1} {side}边距: {actual_mm:.1f}mm，应为{expected_mm}mm')
    
    return issues

def check_line_spacing(docx_path, expected_pt=28):
    """检查行距"""
    doc = Document(docx_path)
    issues = []
    
    for i, para in enumerate(doc.paragraphs):
        if para.paragraph_format.line_spacing_rule == WD_LINE_SPACING.EXACTLY:
            actual_pt = para.paragraph_format.line_spacing / 12700  # EMU to pt
            if abs(actual_pt - expected_pt) > 0.5:
                issues.append(f'段落{i+1}: 行距{actual_pt:.1f}pt，应为{expected_pt}pt')
    
    return issues
```

### 审核报告模板

```markdown
# 投标文件审核报告

## 总体评价
- 完整性：XX/XX 项已响应（XX%）
- 合规性：XX 项实质要求全部满足
- 一致性：XX 处数据不一致
- 格式合规：XX 项格式问题
- 页面布局：XX 项布局问题

## 问题清单
| 序号 | 问题类型 | 位置 | 问题描述 | 严重程度 | 建议修改 |
|------|---------|------|---------|---------|----------|
| 1 | 遗漏 | 第2.3节 | 缺少实施甘特图 | 高 | 补充实施计划甘特图 |
| 2 | 不一致 | 报价表vs技术方案 | 工期描述不一致 | 中 | 统一为6个月 |
| 3 | 格式 | 正文段落 | 字体为宋体，应为仿宋 | 高 | 修改字体 |
| 4 | 页码 | 第15页 | 页码跳号（14→16） | 高 | 检查分页符 |
```

---

## Phase 5: 输出交付 (Deliver)

输出为Markdown或Word格式，包含：
- 完整的投标文件正文
- 招标要求对照表
- 偏离表
- 审核报告

Word输出使用 `word-doc-processing` skill 的 python-docx 方法。

---

## Step 4.2: 页面布局验证（LibreOffice + pdfplumber + PyMuPDF）

### Step 4.2.1: docx → PDF 转换

```bash
# LibreOffice headless 高保真转换
soffice --headless --convert-to pdf --outdir /tmp/ out.docx
```

### Step 4.2.2: 页码连续性检查

```python
import pdfplumber, re

def check_page_numbers(pdf_path, expected_total=None):
    issues = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        if expected_total and total != expected_total:
            issues.append(f'总页数{total}，期望{expected_total}')
        for i, page in enumerate(pdf.pages):
            footer = page.crop((0, page.height - 56, page.width, page.height))
            text = footer.extract_text() or ''
            m = re.search(r'(?:第\s*)?(\d+)', text)
            if m and int(m.group(1)) != i + 1:
                issues.append(f'第{i+1}页: 页脚显示第{m.group(1)}页')
    return issues
```

### Step 4.2.3: 页眉页脚一致性检查

```python
import pdfplumber

def check_header_footer(pdf_path, skip_first=True):
    headers, issues = [], []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if skip_first and i == 0: continue
            header = page.crop((0, 0, page.width, 42))
            headers.append((header.extract_text() or '').strip())
    unique = set(h for h in headers if h)
    if len(unique) > 1:
        issues.append(f'页眉不一致，发现{len(unique)}种')
    return issues
```

### Step 4.2.4: 视觉对比（可选）

```python
import fitz
from PIL import Image
import io

def render_pages(pdf_path, out_dir='/tmp/', dpi=150):
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        pix = doc[i].get_pixmap(dpi=dpi)
        Image.open(io.BytesIO(pix.tobytes('png'))).save(f'{out_dir}/page_{i+1}.png')
```

---

## 相关开源项目

| 项目 | Stars | 功能 | 链接 |
|------|-------|------|------|
| bid-analysis | 24 | 100页招标10分钟解析，3层索引+9个并行提取模块 | github.com/sanwudazhiyuan/bid-analysis |
| BidMaster-Pro | MIT | 全流程：解析→生成→审核→排版，21项合规检查 | github.com/guangshu100/BidMaster-Pro |
| BiaoShu-SKILL | 109 | 14步自动化管线，多行业支持 | github.com/Get00/BiaoShu-SKILL |
| OpenBidKit_Yibiao | — | 18项解析，桌面应用 | github.com/FB208/OpenBidKit_Yibiao |
| yibiao-simple | — | Python+React，LLM提取评分标准 | github.com/yibiaoai/yibiao-simple |

**三层章节识别法**（参考 bid-analysis）：
1. **目录检测**（高置信）：识别目录样式段落，建立章节地图
2. **层级编号**（中置信）：正则匹配 `第X章`、`一、`、`1.1`、`1.1.1` 等
3. **关键词匹配**（低置信）：`资格/资质/评分/技术要求/商务条款/废标` 等

---

## Pitfalls

1. **扫描件PDF**：必须用marker-pdf做OCR，pymupdf只能提取文本PDF
2. **表格提取**：PDF表格可能不完整，用pdfplumber比pymupdf效果好，需人工复核关键数据
3. **实质性要求**：区分「实质性」（不满足即废标）和「一般性」（允许偏离），优先保证实质性
4. **数据来源**：所有事实性数据必须来自用户素材，AI不得编造
5. **格式要求**：严格按招标文件执行（字体、页码、装订等）
6. **前后一致**：工期、人员、报价在各章节间保持一致
7. **废标条款**：特别注意废标/否决投标条款，这些是红线
8. **时间节点**：投标截止、答疑截止、保证金缴纳等关键节点必须提醒用户
9. **保密意识**：不将一个客户的投标内容泄露给另一个客户
10. **章节结构**：不同招标文件结构差异大，必须按实际文件结构来，不套固定模板
11. **模板残留**：技术标从旧项目复制时，经常遗留旧项目名称、地点、参数 — 这是最常见的低级错误
12. **补遗遗漏**：投标期间发的澄清/修改通知必须纳入，否则响应可能过时
13. **业绩三件套**：完整业绩证明 = 中标通知书 + 合同协议书 + 竣工验收证明/业主评价，缺一不可
14. **报价算术**：单价×数量必须验算，汇总表与明细表必须一致
15. **串通投标**：同一IP、同一人、同一保证金账户会被判定串通，注意合规
16. **中文字体陷阱**：python-docx的font.name只返回拉丁字体(w:ascii)，中文字体在w:rFonts/@w:eastAsia中，必须用lxml直接读XML才能验证
17. **字号单位**：python-docx用EMU(1pt=12700EMU)，XML用半磅值(12pt=24)，招标文件用中文字号(小四=12pt)，转换时注意单位
18. **docx→PDF保真度**：LibreOffice转换可能有细微差异（特别是复杂表格），关键页面建议人工比对
19. **国内网络**：GitHub项目可通过镜像站或PyPI镜像安装，详见下方安装指南
