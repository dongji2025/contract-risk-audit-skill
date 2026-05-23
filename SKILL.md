---
name: contract-risk-audit
description: |
  Comprehensive contract and agreement risk analysis tool based on Chinese Civil Code and
  Commercial Secrets Protection Regulations. Reviews uploaded legal documents for potential risks,
  unfair clauses, and concerning terms with clause-by-clause analysis. Use when:
  - User uploads a contract, agreement, or legal document for review
  - User asks to check contract risks or analyze agreement terms
  - User wants to identify red flags or concerning clauses in contracts
  - User requests contract risk assessment or evaluation
  - User needs help understanding legal document implications and risks
  - Analyzing: NDA, employment contracts, service agreements, lease agreements,
    partnership agreements, software licenses, or any legal contracts
  - User asks: "帮我审查这份合同的风险", "分析一下这个协议有什么坑",
    "检查这份合同是否符合《民法典》规定", "商业秘密保护有什么风险" or similar requests
---

# Contract Risk Audit Skill - 合同风险审查技能

## Overview

This skill provides comprehensive contract risk analysis based on **《中华人民共和国民法典》**
and **《商业秘密保护规定》**. It identifies problematic clauses, evaluates risk levels at the
**clause-by-clause** granularity, and provides actionable recommendations with specific legal references.

## Features

- ✅ **《民法典》第463-647条逐条比对** — 系统覆盖合同编全部185条条文，逐一检测风险
- ✅ **《商业秘密保护规定》专项审查** — 保密条款、竞业限制、保密措施等全方位审查
- ✅ **160+风险检测模式** — 基于正则表达式的条款级风险自动识别（含非竞争/竞业限制超期专项检测）
- ✅ **公司主体信息提取** — 自动提取公司名称、信用代码、注册资本、银行账号等
- ✅ **法律依据精准标注** — 每条风险均标注对应的法条编号和完整条文内容
- ✅ **详细修改建议** — 每条风险都附带具体的条款修改建议，无数量限制
- ✅ **图文并茂Word报告** — 专业封面、目录、饼图、仪表盘、柱状图、排名图等完整可视化
- ✅ **风险分级(CRITICAL/HIGH/MEDIUM/LOW)** — 科学的四级风险评估体系
- ✅ **全条款覆盖** — 高/中/低风险条款全部纳入分析，低风险条款含摘要
- ✅ **主体信息核验清单** — 含统一社会信用代码、经营范围、授权代表等核验项

## Quick Start

### Step 1: Extract Text from Contract
```bash
python scripts/extract_text.py contract.docx -o extracted.txt
```

### Step 2: Run Enhanced Clause Analysis
```bash
python scripts/enhanced_clause_parser.py extracted.txt -o enhanced_analysis.json
```

### Step 3: Generate Reports
**Word Report:**
```bash
python scripts/enhanced_word_report.py enhanced_analysis.json -o report.docx
```

**HTML Visual Report:**
```bash
python scripts/enhanced_html_report.py enhanced_analysis.json -o report.html
```

## Core Analysis Process

### 1. Document Parsing (scripts/extract_text.py)

**Supported Formats:**
- PDF (.pdf) - Use pdfplumber or PyPDF2
- Word (.docx) - Use python-docx
- Text (.txt) - Direct reading

### 2. Company Information Extraction

The enhanced parser automatically extracts:
- Company name (公司名称)
- Unified social credit code (统一社会信用代码)
- Legal representative (法定代表人)
- Registered capital (注册资本)
- Paid capital (实缴资本)
- Registered address (注册地址)
- Bank name (开户银行)
- Bank account number (银行账号)
- Contact phone (联系电话)

### 3. Clause-by-Clause Risk Analysis with Full Civil Code Coverage

**逐条比对分析基于以下法律体系：**

#### 《中华人民共和国民法典》合同编 第463-647条 全覆盖：

| 章节 | 条款范围 | 核心审查内容 | 风险检测模式数 |
|------|---------|------------|-------------|
| 第一章 一般规定 | 第463-468条 | 合同定义、约束力、解释规则 | 3 |
| 第二章 合同的订立 | 第469-501条 | 要约承诺、格式条款、预约合同、缔约过失 | 12 |
| 第三章 合同的效力 | 第502-508条 | 生效条件、无权代理、免责条款无效 | 6 |
| 第四章 合同的履行 | 第509-534条 | 全面履行、抗辩权、情势变更、连带责任 | 12 |
| 第五章 合同的保全 | 第535-542条 | 代位权、撤销权 | 3 |
| 第六章 合同的变更和转让 | 第543-556条 | 单方变更禁止、债权转让、债务转移 | 8 |
| 第七章 权利义务终止 | 第557-576条 | 合同解除、不可抗力、后合同义务 | 8 |
| 第八章 违约责任 | 第577-594条 | 违约金、定金、赔偿范围、减损义务 | 10 |
| 第九章 买卖合同 | 第595-630条 | 风险转移、检验验收、权利瑕疵、知识产权 | 10 |
| 第十章 供用电水气热力合同 | 第631-647条 | 公共服务、政府定价 | 3 |

#### 《商业秘密保护规定》专项审查：

| 审查维度 | 法定标准 | 风险检测模式数 |
|---------|---------|-------------|
| 商业秘密三要素 | 秘密性、价值性、保密性 | 5 |
| 保密措施合理性 | 6类法定措施清单 | 4 |
| 保密期限合理性 | 不得超过合理商业寿命 | 5 |
| 竞业限制合规性 | ≤2年 + 经济补偿 | 8 |
| 侵权行为全覆盖 | 盗窃、贿赂、欺诈等 | 3 |
| 违约赔偿责任 | 损失计算、惩罚性赔偿 | 4 |
| 泄密后义务 | 通知、返还、销毁 | 3 |
| **合计** | — | **40+ 专用模式** |

### 4. Risk Detection Patterns

**HIGH Risk (Score 20-25):**
- Format clauses without proper notice (Art. 496 Civil Code)
- One-sided termination rights
- Excessive liquidated damages (>30% of loss)
- Deposits exceeding 20% (Art. 586 Civil Code)
- Unlimited liability clauses
- Permanent confidentiality requirements
- Non-compete exceeding 2 years
- Non-compete without compensation

**MEDIUM Risk (Score 10-15):**
- Ambiguous payment terms
- Unclear change-of-account procedures
- Restrictive transfer clauses
- Broad IP assignment terms
- Vague force majeure definitions

**LOW Risk (Score 5):**
- Standard industry terms
- Normal warranty periods
- Common dispute resolution clauses

### 5. Report Generation

**Word Document Features（图文并茂报告，共10个章节）:**

| 章节 | 内容 | 特色 |
|------|------|------|
| 封面 | 报告标题、风险等级徽章、关键统计数据 | 装饰线框、分类标识、法律依据标注 |
| 目录 | 完整章节目录 | 10个章节的结构化导航 |
| 一、执行摘要 | 关键指标统计表、总体评估 | 风险等级彩色标识、警告提示 |
| 二、风险可视化分析 | 饼图、仪表盘、堆叠柱状图、横向排名图 | 4种图表类型，160 DPI高清 |
| 三、合同主体信息 | 公司基本信息表、主体信息核验清单 | 缺失字段红色标注、5项核验清单 |
| 四、风险概览 | 风险类别分布矩阵 | 6列完整统计表 |
| 五、商业秘密专项审查 | 9项合规检查清单、审查发现、法律依据 | 3列表格（项目/状态/评估）、彩色评估 |
| 六、条款详细分析 | 高风险&中风险逐条深度分析 + 低风险摘要 | 条款原文完整展示、风险识别、法律依据、修改建议 |
| 七、《民法典》覆盖清单 | 法条-条款关联表 | 4列索引表（法条/涉及数/最高风险/相关条款） |
| 八、优先修订建议 | 三级优先级：必须修订/建议协商/通用建议 | 颜色分级、具体修改方案 |
| 九、法律依据索引 | 按法律分组索引 + 相关法规列表 | 结构化引用 |
| 十、免责声明 | 法律效力说明 | 灰色文字标注 |

**报告质量保障：**
- 条款内容全部完整展示（不截断），确保分析完整性
- 风险因素、法律依据、修改建议全部列出（无数量限制）
- 自动检测常见分析缺陷（如竞业限制超期、非竞争条款等）
- 数据源键名自动校验，防止数据丢失

**HTML Report Features:**
- Interactive risk charts (Chart.js)
- Tabbed clause navigation
- Color-coded risk levels
- Legal reference links
- Mobile-responsive design

## Scripts Reference

| Script | Purpose |
|--------|---------|
| `extract_text.py` | Extract text from PDF/DOCX/TXT |
| `enhanced_clause_parser.py` | Clause-level analysis with legal references |
| `enhanced_word_report.py` | Generate detailed Word report |
| `enhanced_html_report.py` | Generate interactive HTML report |

## References

- `references/red_flags.md` - High-risk pattern database
- `references/risk_categories.md` - Risk level definitions
- `references/common_clauses.md` - Standard clause patterns
- `references/civil_code_risks.md` - **《民法典》第463-647条全文风险审查依据**
- `references/trade_secret_regulation.md` - **最新《商业秘密保护规定》全文审查依据**
- `references/industry_standards.md` - Industry benchmarks

## Usage Examples

### Example 1: Full Service Contract Review
```
User: 帮我审查这份服务合同的风险

→ python scripts/extract_text.py service_contract.docx -o extracted.txt
→ python scripts/enhanced_clause_parser.py extracted.txt -o analysis.json
→ python scripts/enhanced_word_report.py analysis.json -o risk_report.docx
→ python scripts/enhanced_html_report.py analysis.json -o risk_report.html
```

### Example 2: Quick Risk Scan
```
User: 快速扫描这份合同中的高风险条款

→ python scripts/extract_text.py contract.docx
→ python scripts/enhanced_clause_parser.py extracted.txt
→ Review HIGH risk items in output JSON
```

### Example 3: Commercial Secrets Focus
```
User: 这份合同中商业秘密保护有什么风险

→ python scripts/enhanced_clause_parser.py extracted.txt -o analysis.json
→ Focus on: confidentiality duration, non-compete, compensation clauses
→ Reference: civil_code_risks.md section on commercial secrets
```

## Legal Disclaimer

This skill provides automated risk assessment based on pattern matching and general
interpretations of Chinese law. It does NOT constitute legal advice. For specific
legal matters, consult a qualified attorney licensed in China.

Generated reports should be reviewed by legal professionals before making
contractual decisions.
