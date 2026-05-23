## 一、合同风险审查skill介绍
      这个skill仅是商业买卖合同的风险审查，适用于国内标的买卖，法律引用还不够完善（本人不是法律专业人员），但是该skill针对每一个存在问题的合同条款，都有详细的法律依据来支撑。
      ### 1.contract-risk-audit-skill技能架构
├── SKILL.md                    # 技能定义文档（核心入口）
├── requirements.txt            # Python依赖配置
├── analysis_result.json         # 分析结果输出
│
├── assets/                      # 资源文件
│   └── risk_report_template.md  # 风险报告模板
│
├── references/                  # 参考资料
│   └── civil_code_risks.md      # 民法典风险条款参考
│
└── scripts/                     # 核心分析脚本
    ├── extract_text.py          # 文本提取模块（PDF/DOCX/TXT）
    ├── enhanced_clause_parser.py # 增强条款解析器（风险分析核心）
    ├── enhanced_word_report.py  # Word报告生成器
    └── __pycache__/             # Python缓存

    ### 2.数据流向
       ┌──────────────────┐     ┌──────────────────────┐     ┌─────────────────┐

│ 合同文件      │ ──→ │ extract_text.py  │ ──→ │ enhanced_clause_     │ ──→ │ enhanced_word_   │

│ .docx/.pdf   │     │ 文本提取          │     │ parser.py            │     │ report.py        │

│ /扫描版PDF   │     │                  │     │ 条款级风险分析        │     │ Word报告生成      │

└──────────────┘     └──────────────────┘     └──────────────────────┘     

                            │                          │                          │

                            ▼                          ▼                          ▼

                     extracted.txt              analysis_result.json       审查报告.docx

                     (纯文本)                    (结构化JSON)               (图文并茂)

    ### 3.运行环境
要有python环境，推荐3.12版本及以上和安装其它依赖包、放在openclaw、claude code、trae等skills目录下，然后调用即可。

## 二、合同风险审查报告结果示例（部分截图）
<img width="1080" height="1143" alt="图片" src="https://github.com/user-attachments/assets/0fe0fc8e-4407-477a-b9ad-e47b84cde810" />
<img width="1080" height="735" alt="图片" src="https://github.com/user-attachments/assets/e9c80a9a-46ae-4a03-be1f-52938abd5511" />
<img width="1080" height="790" alt="图片" src="https://github.com/user-attachments/assets/aae3e1f0-309b-410d-b448-1c4cd26657dc" />
<img width="1080" height="560" alt="图片" src="https://github.com/user-attachments/assets/7d014227-11d0-42f6-b04f-79bc93e19f49" />
<img width="1080" height="930" alt="图片" src="https://github.com/user-attachments/assets/47bef55f-b73a-47f4-a704-b962eeb3cbad" />

## 三、整体效果及后续
     通过该合同风险审查skill还是能发现合同存在的风险的，整体效果见上述示例，报告看上去还是不错的，建议大家可以去github下载测试。该技能测试还不够全面的，希望大家可以在此基础上继续完善。

## 四、我的微信公众号
微信公众号：cnsecux 





