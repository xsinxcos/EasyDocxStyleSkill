# EasyDocxStyleSkill

一个可分享的 Codex Skill 仓库，用于从论文模板 `.docx` 中提取格式规则，生成可持久化的规则包，并基于规则包对论文进行严格审查与修复。

当前仓库提供的核心 skill 是：

- `thesis-format-enforcer`

它的设计目标不是“看起来差不多”，而是：

- 先从模板中抽取规则
- 再把规则持久化为 rule pack
- 再用 rule pack 反向重建模板
- 最后从 DOCX 包元数据层面对比原模板与重建结果

只有差异为 `0` 时，规则包才被视为有效。

## 主要能力

- 从学校/学院/期刊模板 `.docx` 生成可复用规则包
- 同时吸收两类规则：
  - 模板中可直接测量的格式
  - 模板正文中写出来的格式要求
- 用规则包重建模板并做 DOCX package metadata 校验
- 对论文草稿执行审查，输出 JSON 和 Markdown 报告
- 修复前强制先审查，再显式征得用户同意
- 通过 Codex hooks 约束流程顺序，避免跳步

## 仓库结构

```text
.
├─ .agents/
│  └─ skills/
│     └─ thesis-format-enforcer/
│        ├─ SKILL.md
│        ├─ agents/openai.yaml
│        ├─ references/
│        │  ├─ workflow.md
│        │  └─ rule-pack-schema.md
│        └─ scripts/
│           ├─ docx_model.py
│           └─ thesis_format_pipeline.py
├─ .codex/
│  ├─ hooks.json
│  └─ hooks/
│     ├─ user_prompt_submit.py
│     ├─ pre_tool_use.py
│     └─ stop_guard.py
├─ requirements.txt
└─ README.md
```

## 环境要求

- Python 3.10+
- Codex / 兼容 skill 的本地运行环境
- Windows、macOS、Linux 均可

安装 Python 依赖：

```bash
pip install -r requirements.txt
```

## Skill 核心流程

### 1. 构建规则包

```bash
python .agents/skills/thesis-format-enforcer/scripts/thesis_format_pipeline.py build-rule-pack --template path/to/template.docx --name school-2026
```

会生成：

- `rules.json`
- `manifest.json`
- `text-requirements.md`
- `package-manifest.json`
- `validation/validation-report.json`
- `validation/validation-report.md`

### 2. 校验规则包

```bash
python .agents/skills/thesis-format-enforcer/scripts/thesis_format_pipeline.py validate-rule-pack --rule-pack artifacts/rule-packs/school-2026 --template path/to/template.docx
```

规则包必须通过 DOCX 包级元数据校验后，才能继续用于审查或修复。

### 3. 审查论文

```bash
python .agents/skills/thesis-format-enforcer/scripts/thesis_format_pipeline.py audit --document path/to/thesis.docx --rule-pack artifacts/rule-packs/school-2026
```

会输出：

- `audit-report.json`
- `audit-report.md`

报告会尽量给出问题位置和文本范围，例如：

- 段落号
- 页眉位置
- 表格单元格位置
- 文本前 18 字和后 18 字

### 4. 修复论文

```bash
python .agents/skills/thesis-format-enforcer/scripts/thesis_format_pipeline.py repair --document path/to/thesis.docx --rule-pack artifacts/rule-packs/school-2026 --output path/to/fixed.docx --approval-id <token>
```

注意：

- 不能跳过审查直接修复
- 必须先询问用户
- 必须在后续 turn 中获得明确批准
- `approval-id` 由 hook 机制配合工作流生成

## 模板文字规则支持情况

这个 skill 不只看模板样式，还会看模板正文里写出来的格式要求。

当前已支持落地为可执行规则的典型目标包括：

- 中文摘要正文
- 英文摘要正文
- 一级标题
- 二级标题
- 三级标题
- 正文段落
- 页眉文本
- 表题
- 表格正文
- 图题
- 代码块
- 参考文献条目

仍可能保留为 `stored_only` 的目标包括：

- 仅能做人工判断的一致性要求
- 难以稳定映射到确定 DOCX 对象的抽象规则

## Hooks 说明

仓库内置了 Codex hooks，用来保证流程按预期推进：

- `UserPromptSubmit`
  - 注入论文格式工作流上下文
- `PreToolUse`
  - 阻止跳过规则包校验、跳过审查直接修复等违规步骤
- `Stop`
  - 要求输出中明确指出生成的工件路径

对应文件：

- `.codex/hooks.json`
- `.codex/hooks/user_prompt_submit.py`
- `.codex/hooks/pre_tool_use.py`
- `.codex/hooks/stop_guard.py`

如果你只复制 skill，而不复制 `.codex/`，那么流程约束不会自动生效。

## 分享到 GitHub 的建议

建议只提交代码和文档，不提交以下内容：

- 真实论文
- 学校模板原件
- `artifacts/` 运行产物
- 本地 IDE 配置

本仓库的 `.gitignore` 已默认忽略：

- `artifacts/`
- `.codex/state/`
- `.idea/`
- `*.docx`

如果你想公开一个最小可运行仓库，建议保留：

- `.agents/skills/thesis-format-enforcer/`
- `.codex/`
- `requirements.txt`
- `README.md`

## 推荐发布方式

### 方式一：直接作为仓库分享

适合让别人 clone 后本地使用。

### 方式二：只发布 skill 目录

如果接收方已经有自己的 Codex 工程，可以只复制：

```text
.agents/skills/thesis-format-enforcer/
```

如果需要同样的流程保护，再额外复制：

```text
.codex/
```

## 已知限制

- 不能保证分页结果与模板视觉上 100% 一致
- 浮动对象、复杂图文混排、部分域对象仍可能需要人工复核
- 某些“统一风格”类要求目前仍更适合人工判断

## 后续可继续扩展的方向

- 进一步降低审查噪音
- 增加图题/表题相对位置校验
- 增加强制引用关系校验
- 增加更多高校模板的适配样例
- 把依赖和 CLI 再包装成可安装工具

## 说明

仓库当前未附带开源许可证。

如果你准备公开分享到 GitHub，建议你根据自己的分享范围选择并补充合适的许可证，例如：

- MIT
- Apache-2.0
- GPL-3.0
