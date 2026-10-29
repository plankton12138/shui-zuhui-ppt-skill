---
name: shui-zuhui
description: 水组会——把一篇论文（PDF）做成组会/文献汇报用的极简 PPT，并附逐页讲解稿。适用于"论文转 PPT""组会汇报""文献分享""paper presentation""读论文讲给别人听""水组会"等请求。产出为白底、无装饰、红字标重点、含原论文图表的中文 PPT。
---

# 水组会 · 论文 → 组会 PPT

## 宗旨

**不想做论文汇报 PPT 了，想水一篇，又能让别人以为是认真做的。**

这是所有取舍的最终判据，一切规则都是它的具体化。两头都要占住：

- **投入小（制作端）**：能自动化的全部交给脚本（抽图、排版、渲染校验）；文案按叙事骨架填空，不逐段精读论文；手工时间只花在两处——裁图干不干净、标题像不像论断。遇到要花大量时间且听众看不出来的环节（重绘图、逐条翻译 related work、做动画），直接放弃
- **观感足（观看端）**："认真做过"的信号一个都不能少：原论文的图无损入页、每页标题是一句论断、重点红字带机理解释、附一份能照着讲的逐页讲稿。这些信号恰恰都由脚本代劳，不额外花你的时间
- **验收标准**：听众觉得"这个人把论文吃透了"，而实际投入远低于认真做一份汇报的耗时

把一篇论文压成 8–12 页"只有干货"的组会 PPT，并给出一份可以直接照着讲的口播稿。目标不是"把论文抄一遍"，而是让听众觉得**这个人把论文吃透了**。

## 版式契约（不可协商）

本 skill 的版式基调是固定的，不在单次任务里调整。落到执行上：

- **纯白背景、黑字**，强调色克制：默认用纯红 `#FF0000` 点睛；确实需要区分两类信息时可以引入第二种强调色（如蓝 `#0070C0`，用 `<b>`/`<bb>` 标记），但全篇强调色不超过两种
- **不要**目录页、致谢页、过渡页、页码、logo、图标、色块、动画、渐变
- **每页一个论断**，标题就是这句论断（不要写"研究方法"四个字就完事，写"研究方法：四个关键设计"）
- **必须出现原论文的图或表**，从 PDF 里无损裁出来，不要重绘、不要截图带页眉
- **不要变量级细节**：不抄公式、不列符号表、不逐条翻译 related work
- 目标是让听众觉得"这个人把论文吃透了"，而不是"这个人把论文抄了一遍"

## 页数与信息量

叙事骨架是参考，不是硬指标。**页数以把内容讲清为准，通常 8–12 页，可按论文体量增减**：方法复杂就多拆一页，某部分确实单薄就并掉，不凑页数也不硬压。判断标准是每一页都有信息量——如果一页删掉后听众什么都没少，就不该有这一页。

## 叙事骨架（参考，9 页起步）

| # | 页面      | 必须回答的问题                                        |
| - | ------- | ---------------------------------------------- |
| 1 | 封面      | 题目、出处、一句话中文题眼                                  |
| 2 | 研究背景    | 研究对象是什么？为什么现在做？卡在哪两个挑战上？                       |
| 3 | 研究现状    | 有哪几条技术路线？每条路线的代表工作 + **一句红字缺陷** → 共同局限 = 本文立题点 |
| 4 | 研究方法（一） | 总体架构长什么样（配架构图）                                 |
| 5 | 研究方法（二） | 关键设计"为什么这么设计"，每个设计对应背景里的一个障碍                   |
| 6 | 实验设置    | 测试床、负载、指标、场景（配测试床表/图）                          |
| 7 | 实验结果（一） | 主结果：数字 + 一句机理解释                                |
| 8 | 实验结果（二） | 补充结果 / 开销 / 鲁棒性                                |
| 9 | 结论与不足   | 三条结论 + 四到五条不足（含自己加的评估视角）                       |

页数按论文体量在 8–12 之间浮动，有充足理由也可以更多或更少；方法复杂就拆两页，实验单薄就并一页。**不要**为了凑页数加水，也不要为省页数砍掉信息量。

写作细则见 `references/story-guide.md`，版式细则见 `references/style-spec.md`，  
口播稿写法见 `references/narration-guide.md`。开工前先读 story-guide。

## 工作流

### 1. 摸清 PDF

```bash
python scripts/extract_figures.py paper.pdf list          # 有文字层：列出图/表标题
python scripts/extract_figures.py paper.pdf text --pages 1-4
```

如果 `list` 报 `NO CAPTIONS FOUND`，说明 PDF 没有文字层（字体被转成曲线，常见于 arXiv 某些模板）。  
此时**不要**上 OCR，直接把页面渲成图用眼睛读：

```bash
python scripts/extract_figures.py paper.pdf page --pages 1-12 --dpi 150 --outdir pages
```

然后逐页 `Read` 这些 PNG。读的时候只记四件事：问题、现状的缺陷、方法的三个设计、实验的三个数字。

### 2. 抽图

有文字层时优先自动裁：

```bash
python scripts/extract_figures.py paper.pdf auto --caption "Fig. 2" --out figs/fig2.png --dpi 400
```

无文字层时用相对坐标裁，**裁完必须 `Read` 看一眼**：

```bash
python scripts/extract_figures.py paper.pdf crop --page 9 --relative 0.21,0.13,0.79,0.26 --out figs/fig2.png --dpi 400
```

裁图三条铁律：

1. 图注（caption）如果自带英文，就裁掉，改用 11pt 灰色中文图注；否则会中英重复
2. 页眉（running head，如 `Resource Manager 7`）必须裁掉，否则图上会有一条脏边
3. 宽表（aspect > 3）不要硬塞进半栏，用 `figure` 类型的宽图布局

### 3. 写 deck.json

内容全部放在一个 JSON 里，脚本只负责排版。**先写文案，再调版式**。  
完整字段说明见 `references/style-spec.md`，行内标记与版式常量都以它为准。

```json
{
  "title_en": "...", "title_zh": "...", "venue": "arXiv:xxxx · 2025.03", "authors": "...",
  "slides": [
    {"type": "cover"},
    {"type": "points", "title": "研究背景", "lead": "……<r>关键概念</r>", "bullets": ["**背景**：……"]},
    {"type": "columns", "title": "研究现状", "columns": [{"head": "云侧", "lines": ["..."], "tail": "缺陷：……"}], "footer": "**共同局限**：……"},
    {"type": "figure", "title": "研究方法（一）：总体架构", "image": "figs/fig1.png", "caption": "Fig. 1  架构设计（原论文）", "bullets": ["..."]},
    {"type": "figure_stack", "title": "...", "images": [{"path": "figs/fig2.png", "caption": "..."}], "bullets": ["..."]},
    {"type": "figure_pair", "title": "...", "images": [{"path": "..."}, {"path": "..."}], "bullets": ["..."]},
    {"type": "two_col", "title": "结论与不足", "left": {"head": "结论", "bullets": ["..."]}, "right": {"head": "不足", "bullets": ["..."]}}
  ]
}
```

行内标记只有五种：`**加粗**`（小标题/术语）、`<r>红字</r>`（重点）、`<rb>红字加粗</rb>`、`<b>蓝字</b>`（第二强调色，用于与红区分的同类信息）、`<bb>蓝字加粗</bb>`。

### 4. 生成 + 自查

```bash
python scripts/build_deck.py deck.json -o out.pptx
python scripts/render_deck.py out.pptx --outdir preview     # Windows + PowerPoint
```

**必须**把 preview 里的每一页都 `Read` 一遍再交付。检查清单：

- [ ] 标题没有折行（折行会压到正文；脚本会自动缩字号，但 26pt 仍折行说明标题该改短）
- [ ] 没有文字溢出边框、没有文字被图片压住
- [ ] 每页强调色（红+蓝合计）不超过 3 处，且强调在"缺陷 / 机制 / 关键数字"上
- [ ] 图里没有页眉、没有残留的英文图注
- [ ] 图注紧贴图片下沿，不是飘在页面底部
- [ ] 黑灰之外的全篇强调色不超过 2 种

### 5. 写讲解稿

`references/narration-guide.md` 给了模板。要点：**每页 100–180 字口播**，先定位（这页在整篇的哪一环），再给结论，最后一句过渡。另外附 4–6 个"可能被追问"的问题和答法。

### 6. 交付

把 pptx、讲稿 md、preview 图一起放到用户指定目录，用 `present_files` 打开。

## 脚本

| 脚本                           | 作用                                         | 依赖                            |
| ---------------------------- | ------------------------------------------ | ----------------------------- |
| `scripts/extract_figures.py` | 列图注 / 导文字 / 渲页面 / 按 caption 自动裁图 / 按相对坐标裁图 | PyMuPDF                       |
| `scripts/build_deck.py`      | JSON → pptx，固定版式 + 自动缩字号防溢出                | python-pptx, Pillow           |
| `scripts/render_deck.py`     | pptx → 逐页 PNG，供肉眼校验                        | PowerPoint COM（或 LibreOffice） |

安装依赖：`pip install -r requirements.txt`（pymupdf / python-pptx / pillow / pywin32）

## 常见坑

- **PDF 无文字层**：`list` 和 `text` 都会空。直接渲图读，别浪费时间装 OCR。
- **自动裁图裁到半栏**：双栏论文里图注在左栏、图跨栏时，`auto` 可能只裁一半。改用 `crop --relative` 手调，`Read` 验证。
- **图裁进了正文**：`auto` 的 padding 默认 4pt，如果图上方紧贴正文，把 `--padding` 调成 0 或负数再试。
- **字号被压到 14pt**：说明这页文字太多。删字，不要缩字号——14pt 在投影上已经不可读。
- **PowerPoint COM 报错**：`render_deck.py` 会退回 LibreOffice；两者都没有就只能靠 `python-pptx` 读文本框尺寸粗查。
