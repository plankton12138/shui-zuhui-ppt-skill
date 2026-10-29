# 水组会 · shui-zuhui

> 一篇论文进去，一盏茶的工夫，一份看起来读了一个礼拜的组会汇报出来——白底、无装饰、红字点睛、原图直入，附逐页讲稿。

---

## 功能

**一句话：把论文 PDF 压成"只有干货"的汇报材料，而不是把论文抄一遍——省时间的是你，看到"认真"的是听众。**

- **固定叙事骨架，页数弹性** —— 研究背景 → 研究现状 → 研究方法 → 实验 → 结论与不足；页数以讲清为准，通常 8–12 页，按论文体量增减
- **零装饰版式** —— 纯白底、黑字、强调色克制（默认红 `#FF0000` 点睛，需要区分同类信息时可加蓝 `#0070C0`，全篇不超过两种强调色）；没有目录页、致谢页、页码、图标、色块、动画
- **原图直入** —— 从 PDF 按 400 dpi 无损裁出论文自己的图与表，不重绘、不带页眉
- **自动排版** —— 标题按字宽自动缩字号保证单行，正文按容器高度自动降字号，图注紧贴图片下沿
- **双产出** —— 一份 `.pptx`，一份逐页讲稿（每页 100–180 字口播 + 常见追问的答法）
- **内容与版式分离** —— 文案写在 `deck.json` 里，换论文只改文案，版式不会漂
- **七种版式** —— 封面 / 纯文字 / 多栏对照 / 单图+侧栏 / 双图叠放 / 双图并排 / 双栏结论
- **兜底能力强** —— 无文字层的 PDF（字体转曲线）也能处理：渲成图读，按坐标裁

---

## 产出

```
out.pptx      9 页左右的汇报稿
讲稿.md       逐页口播 + "如果被追问"的备答 + 时间分配建议
```

典型页面构成（以一篇典型论文为例，页数随内容伸缩）：

| 页   | 内容                          |
| --- | --------------------------- |
| 1   | 封面：英文原题、中文题眼、出处             |
| 2   | 研究背景：对象、三重异构、两个挑战           |
| 3   | 研究现状：三条技术路线，每条一句红字缺陷 → 共同局限 |
| 4–5 | 研究方法：总体架构（配架构图）+ 关键设计的动机    |
| 6   | 实验设置：测试床、负载、指标、场景（配表）       |
| 7–8 | 实验结果：数字 + 机理，配原论文曲线图        |
| 9   | 结论与不足：左 3 条结论，右 4–5 条不足     |

---

## 使用样例

用法只有一种：**对 agent 说一句话**（WorkBuddy / Claude Code 等，把路径换成你自己的）。读论文、裁图、排版、自查全由 agent 和脚本完成，不需要人手动跑任何代码。下面几种说法覆盖了绝大多数场景：

### 样例 1：最省事——丢个 PDF 就行

> 帮我把这篇论文做成组会 PPT：`D:/papers/xxx.pdf`

agent 全自动走完：读论文 → 裁原图 → 写文案 → 生成 pptx → 渲图自查 → 附逐页讲稿。适合"今晚就要，别问我细节"。

### 样例 2：控制篇幅与侧重

> 把 `D:/papers/xxx.pdf` 做成组会汇报，重点讲方法设计，实验并成一页，总共 8 页左右。

页数和详略都听你的；骨架不变，省得自己想结构。

### 样例 3：按场合定制讲稿

> 下周三组会讲这篇，给我 15 分钟的版本：`xxx.pdf`，讲稿按 15 分钟分配时间。

产出除了 pptx 还会带一份逐页口播稿 + 常见追问的备答。

### 样例 4：改一版已有的

> 第三页的现状对照，路线 B 换成 2024 年那篇 follow-up；第七页的图换成 Fig. 6。

内容都在 `deck.json` 里，agent 改完重新生成即可，版式不会乱——改十版也不漂。

### 样例 5：强调色分工（红 + 蓝）

> 讲 `xxx.pdf`，现状页把每条路线的"缺陷"标红、"已有优势"标蓝，其他页正常。

对应到文案里的行内标记：`<r>缺陷</r>`、`<b>优势</b>`。红是默认重点，蓝只在需要区分两类信息时用。

---

## 安装

把下面这句发给 agent（WorkBuddy / Claude Code 等）即可，**其余全部由 agent 自己完成**——下载、装依赖、验证都不需要人动手：

```text
请安装 skill「shui-zuhui」：把仓库 https://github.com/plankton12138/shui-zuhui-ppt-skill 完整获取（git clone 或下载解压均可）到你的 skills 目录，然后在该目录执行 pip install -r requirements.txt，装好后读取 SKILL.md 并告诉我已就绪。
```

agent 安装时会自行处理这些细节：

- **装到哪里** —— skills 目录即安装。WorkBuddy 为 `~/.workbuddy/skills/shui-zuhui`，WorkBuddy AI（CLI）为 `~/.workbuddy-ai/skills/shui-zuhui`，Claude Code 为 `~/.claude/skills/shui-zuhui`，其余兼容 Agent Skills 规范的工具按各自的 skills 目录约定；也可以装到项目级（如 `<项目>/.workbuddy/skills/`）只在单个项目生效
- **装依赖** —— 要求 Python 3.9+，`pip install -r requirements.txt`：pymupdf（读 PDF）、python-pptx（写 PPT）、pillow（量图片尺寸）、pywin32（Windows 下导预览图，可选）
- **验证** —— 读取 `SKILL.md` 确认 frontmatter 完整；之后说"帮我把这篇论文做成组会 PPT"能命中即成功，未命中就重启 agent 会话再试

---

## 原理

```
PDF ──► 抽取内容 ──► deck.json ──► 渲染器 ──► .pptx
                        │                        │
                        └──► 讲稿.md            └──► 逐页 PNG 校验
```

三层分离：

**内容层** `deck.json` —— 只写文案，用 `**加粗**`、`<r>红字</r>` 两个行内标记控制强调。

**版式层** `build_deck.py` —— 版式是常量，不是参数。所有页面走同一套网格（16:9、0.92 in 左边距、40pt 标题、24pt 正文、等线），渲染器负责：等比放图、按容器高度自动降字号、标题单行化、图注定位。

**校验层** `render_deck.py` —— 用 PowerPoint COM 导出逐页 PNG，让人眼确认没有溢出、没有压字、没有脏边。自动排版只能估算，最终验收靠看。

这样设计的好处：文案可以反复改而不影响版式；版式只在一个文件里改，不会出现"这页 40pt 那页 36pt"的漂移；出问题能定位到是哪一层。

---

## deck.json 速查

```json
{
  "title_en": "论文英文全名",
  "title_zh": "中文题眼",
  "venue": "arXiv:2411.16697 [cs.DC] · 2024.11",
  "slides": [
    { "type": "cover" },
    { "type": "points",  "title": "研究背景", "lead": "…<r>关键概念</r>", "bullets": ["**背景**：…"] },
    { "type": "columns", "title": "研究现状", "columns": [{"head":"云侧","lines":["…"],"tail":"缺陷：…"}], "footer": "**共同局限**：…" },
    { "type": "figure",  "title": "研究方法（一）：总体架构", "image": "figs/fig1.png", "caption": "Fig. 1  …", "bullets": ["…"] },
    { "type": "figure_stack", "images": [{"path":"…","caption":"…"}], "bullets": ["…"] },
    { "type": "figure_pair",  "images": [{"path":"…"},{"path":"…"}], "bullets": ["…"] },
    { "type": "two_col", "title": "结论与不足", "left": {"head":"结论","bullets":["…"]}, "right": {"head":"不足","bullets":["…"]} }
  ]
}
```

行内标记：`**加粗**`、`<r>红字</r>`、`<rb>红字加粗</rb>`、`<b>蓝字</b>`、`<bb>蓝字加粗</bb>`（蓝为可选第二强调色）。  
全部字段与版式常量见 [`references/style-spec.md`](references/style-spec.md)。

---

## 目录结构

```
shui-zuhui/
├── SKILL.md                      # 技能入口：版式契约 + 工作流 + 自查清单
├── README.md                     # 本文件
├── requirements.txt
├── references/
│   ├── style-spec.md             # 画布/字号/颜色/字体 + deck.json 全字段
│   ├── story-guide.md            # 9 页叙事骨架、红字规则、长度控制
│   └── narration-guide.md        # 口播稿写法、备用问答
├── scripts/
│   ├── extract_figures.py        # list / text / page / auto / crop / images
│   ├── build_deck.py             # deck.json → pptx
│   └── render_deck.py            # pptx → PNG（PowerPoint COM / LibreOffice）
```

---

## 适用与不适用

**适合**：计算机方向的会议/期刊论文，有架构图或实验曲线，需要在组会、文献分享、开题汇报上讲 10–15 分钟。

**不适合**：纯理论论文（没有图可放）、综述（页数会超）、需要逐式推导的场合——它的目标是"讲清问题和思路"，不是"复现推导"。

---

## 已知限制

- **无文字层 PDF**：`list` / `text` 会全空，只能渲图肉眼读。不要指望 OCR——公式和图注 OCR 出来也是错的。
- **双栏跨栏图**：自动裁切可能只取到半栏，改用 `crop --relative` 手调。
- **模板页眉**：Springer 等模板每页有 running head，会跟着图一起被裁进来，需要按 pt 精确切边。裁完必须看一眼。
- **字号触底 14pt**：说明这页字太多，应当删字而不是继续缩——14pt 在投影上已不可读。
- **预览依赖 Office**：没有 PowerPoint 时退回 LibreOffice；两者都没有则无法完成视觉验收。

---

## 示例

完整样例产物不随仓库分发。真实论文的一套标准产物包括：`deck.json` 文案 · `figs/` 裁好的原图 · `out.pptx` 成品 · `讲稿.md` 逐页口播 · `preview/` 逐页校验图。

换论文时让 agent 参照 [`references/style-spec.md`](references/style-spec.md) 与 [`SKILL.md`](SKILL.md) 的工作流从零生成即可，结构都是固定的。
