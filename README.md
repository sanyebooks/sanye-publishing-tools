# 孙三野出版工具

面向出版社编辑、图书营销、发行和新媒体团队的 AI 出版工具集。

当前已发布「图书卖点拆解」：孙三野的图书营销卖点拆解技能。读取 PDF、Word、目录、样章或图书简介，生成有原文依据、页码线索、营销优先级和人群圈层的卖点表，并可输出 Excel。

## 三平台目录

| 平台 | 目录 | 用途 |
|---|---|---|
| Codex | `plugins/book-selling-points/` | Codex Plugin 与 Marketplace |
| WorkBuddy | `workbuddy/plugins/book-selling-points/` | WorkBuddy Plugin |
| SkillHub | `skills/book-selling-points/` | SkillHub 从 GitHub 导入的标准 Skill |

三个版本共享同一套卖点拆解标准，但使用各平台对应的清单和目录结构。

## 核心能力

- 默认生成 30 条有原文依据的营销卖点。
- 标注页码或待核的章节位置线索。
- 区分核心层、扩展层和潜力层读者。
- 标注加星和超级加星的营销优先级。
- 按需生成小红书文案和 AI 配图提示词。
- 在平台具备电子表格能力时交付 `.xlsx`。

## 快速测试

> 使用图书卖点拆解，把这本书拆成 30 条有书内依据的营销卖点，标注页码、星级和人群圈层，并生成 Excel。

## 文档

- [Codex 安装指南](docs/Codex安装指南.md)
- [WorkBuddy 安装指南](docs/WorkBuddy安装指南.md)
- [编辑使用手册](docs/编辑使用手册.md)
- [示例 Excel](examples/示例成品-图书营销卖点表.xlsx)

## 安全与授权

请只处理你有权使用的书稿。宣传物发布前，请由编辑复核人物名、数据、原意和页码。本仓库为公开可安装的社区试用版，不等于放弃知识产权。详见 [COMMUNITY-LICENSE.md](COMMUNITY-LICENSE.md)。

Copyright © 2026 孙三野。
