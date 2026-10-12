# Skills

个人沉淀的 AI Agent 技能（Skill）合集，采用 `SKILL.md` 规范，适用于 Kimi Work 、Workbuddy等支持该规范的环境。

## 技能列表

### video-director-guide

视频 AIGC 创作流程规范与避错指南。强制约束创作流程（创意发散 → 形象确认 → 分镜规划 → 美术基准 → 生成 → 迭代），包含镜头拆分原则、参考图使用规范、修改纪律，以及基于实战的模型能力边界总结和踩坑案例集。

### pptx-to-pdf-windows

Windows 上用本机 PowerPoint COM 把 PPTX 转为 PDF 并压缩体积。记录了 PrintRange 必须显式传 None、含空格长路径报 `0x80070003`、剥离内嵌字体解决"储存错误"等关键坑位。

## 使用方式

将技能目录整体放入 Agent 的 skills 加载目录即可被发现；`SKILL.md` 为入口文件，`references/` 下为参考文档。

## License

[MIT](LICENSE)
