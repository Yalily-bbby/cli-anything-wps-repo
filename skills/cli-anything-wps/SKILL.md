---
name: cli-anything-wps
description: WPS Office 自动化命令行工具。通过 COM 自动化接口操控 WPS 文字（Writer）、表格（Calc）、演示（Impress），支持新建/编辑文档并导出 DOCX、XLSX、PPTX、PDF、CSV、HTML，以及 JSON 数据驱动的 PPT 批量自动生成。当用户需要在 Windows 上程序化生成或修改 Office 文档、尤其是用数据批量生成 PPT 时使用。
---

# cli-anything-wps

通过 COM 自动化接口操控 WPS Office 的命令行工具。支持 WPS 文字（Writer）、WPS 表格（Calc）和 WPS 演示（Impress）。

## 前置条件（本机已满足）

- **Windows 操作系统**
- **WPS Office**（本机已安装 12.1.0.28505）
- **Python 3.10+**
- **pywin32**

## 本机安装信息

| 项目 | 位置 |
|------|------|
| CLI 可执行文件 | `D:\Scripts\cli-anything-wps.exe` |
| 包版本 | cli-anything-wps 1.0.0 |
| 所属解释器 | `D:\python.exe`（Python 3.14，另在助手沙箱 Python 中也已安装） |
| 仓库源码 | `C:\Users\yalily\Doubao\harness-anything\harness-anything-master` |
| 参考案例（16 个 PPT 项目） | `C:\Users\yalily\Doubao\harness-anything\harness-anything-master\WPS` |

验证安装：

```bash
cli-anything-wps --help
```

---
## ⚠️ 本机修复说明（2026-10-08 安装时打的补丁）

上游 1.0.0 在本机（WPS 12.1）存在若干问题，安装时已就地修复。三个副本均已同步：
助手沙箱 Python、`D:\python.exe`、仓库源码 `C:\Users\yalily\Doubao\harness-anything\harness-anything-master`。

| 问题 | 现象 | 修复 |
|------|------|------|
| 演示文稿保存格式常量错误 | 导出的 .pptx 是空壳，一张幻灯片都没有 | `ppSaveAsOpenXMLPresentation` 由 `1` 改为 `24` |
| WPS 演示对象没有 `SaveAs2` | 导出 PPT 报 `AttributeError: Add.SaveAs2` | impress 分支改用 `SaveAs` |
| 新建演示文稿默认 0 张幻灯片 | `doc.Slides(1)` 越界报错 | 失败时回退为 `Slides.Add(1, 2)` |
| 会误退出用户已打开的 WPS | 自动化收尾时 `app.Quit()` 连带关闭用户正在编辑的文档 | 改为仅在「调用前系统中没有任何 WPS 组件进程」时，才认为实例是本工具启动的，才允许退出 |

### ⚠️ 安全注意

- **不要执行**上游文档里那句 `taskkill //F //IM wps.exe //T`——它会强制关闭用户正在编辑的 WPS 文档。补丁生效后正常导出不需要它。
- 用户 WPS 已打开时：工具会复用该实例新建文档，导出后只关闭**本次新建的**文档，不会再退出 WPS 本体。
- 已实测：Writer 导出 .docx ✅、Impress 导出 .pptx（含幻灯片内容）✅、Impress 导出 PDF ✅。

---
## JSON 数据驱动 PPT 自动生成 ⭐ 生产环境核心工作流

**这是最常用的模式**：搜数据 → 生成 matplotlib 图表 → 写 JSON → WPS COM 一键生成。

### 完整流程（6步）

```
1. 提取模板背景: 模版.pptx → template_bg.png (960x540)
2. 搜索数据: 招生分数线/学科排名/招生计划/科研数据
3. matplotlib 生成图表: 柱状图/饼图/折线图/横向柱/气泡图
4. 编写 data.json: 12-15页 elements[] 编排（标题≤4字 + 间距≥24pt）
5. python build_xxx.py: WPS COM 逐页构建
6. 输出: PPTX + PDF 双格式
```

### 标准页序

```
S1 封面 → S2 目录 → S3-S10 内容页(图+表+卡片) → S11 总结卡片 → S12 致谢
```

### 元素类型路由

| type | 用途 |
|------|------|
| `text` | 文本框 |
| `image` | 图片(含matplotlib图表) |
| `table` | 表格 |
| `cards_2x3` | 2行x3列彩色卡片 |
| `cards_1x4_info` | 4列数字统计卡 |
| `card_list_wide` | 目录编号列表 |
| `tagline_bar` | 页面底部总结条 |

### 关键约束 ⚠️

- **标题与内容间距**：标题 y=14 h=40(结束于y≈54)，第一个内容元素**必须**起始于 y≥76-78(≥24pt gap)
- **标题最多4字**，居中，SimHei 40-44pt，品牌色，**无装饰线**
- **JSON中文引号**：文本内引用用「」代替 `""`
- **所有元素不出画布**：960×540，y+h≤518
- **WPS COM**：`Fill.ForeColor.RGB`，`SaveAs(path, 32)`导出PDF
- **执行前清理**：`taskkill //F //IM wps.exe //T`

### 参考案例

`C:\Users\yalily\Doubao\harness-anything\harness-anything-master\WPS` 目录下 16 个项目：

| 项目 | 页数 | 主题 |
|------|------|------|
| 清华协和/兰州大学/同济医学院/哈工大/重庆大学/南华大学 | 12页 | 各校招生 |
| 中山大学/中科大/国科大 | 15页 | 名校招生 |
| 复旦大学 | 12页 | 新工科+医学院 |
| 南科大肿瘤医院 | 12页 | 联合培养硕博 |
| 北大/清华/南科大/华中科大/浙大城市学院 | 9-14页 | 各校介绍 |

```bash
# 完整执行示例
python -c "import zipfile; ..."  # 提取模板
python gen_charts.py              # 生成图表
python -c "import json; json.load(open('data.json','r',encoding='utf-8'))"  # 验证JSON
taskkill //F //IM wps.exe //T
python build_xxx.py               # 构建PPTX+PDF
```

## 动效：过场切换与元素动画（2026-10-08 实测）

### 过场切换 ✅ 可用，推荐

`Slide.SlideShowTransition` 完全可用，设置后会写进 `ppt/slides/slideN.xml` 的 `<p:transition>`，
并且能在 WPS 里正确读回。

```python
tr = slide.SlideShowTransition
tr.EntryEffect = 3852      # 推入
tr.Duration = 0.6          # 秒
```

常用取值（实测保存后 XML 里对应的效果标签）：

| 取值 | 效果 | XML |
|------|------|-----|
| 1793 | 淡入淡出 | `<p:fade thruBlk="1"/>` |
| 3852 | 推入 | `<p:push/>` |
| 2817 | 擦除 | `<p:wipe/>` |
| 1281 | 覆盖 | `<p:cover/>` |
| 1537 | 溶解 | `<p:dissolve/>` |
| 3845 / 3846 | 圆形 / 菱形 | `<p:circle/>` / `<p:diamond/>` |
| 3847 / 3848 | 梳理（横 / 竖） | `<p:comb/>` |
| 3850 / 3851 | 新闻快报 / 加号 | `<p:newsflash/>` / `<p:plus/>` |
| 1025 / 1026 | 棋盘（横 / 竖） | `<p:checker/>` |
| 2305 | 随机线条 | `<p:randomBar/>` |
| 513 | 随机 | `<p:random/>` |
| 257 / 258 | 无 | `<p:cut/>` |

实测写入无效（保存后 transition 里没有效果节点）：1289、1538、2821、5137、2561、2320。

### 元素进入动画 ⚠️ 只有「出现」一种

- **不要用** `Slide.TimeLine.MainSequence.AddEffect()`：它会「假成功」——返回一个 EffectType，
  但 `MainSequence.Count` 恒为 0，继续调用就报 `-2147023170 远程过程调用失败`、
  `-2147023174 RPC 服务器不可用`，COM 直接崩。这条接口在 WPS 里不可依赖。
- 用旧接口 `Shape.AnimationSettings`：

```python
an = shape.AnimationSettings
an.Animate = True
an.EntryEffect = 3844     # 3844 = 出现；写别的值会被归一化成 3844，WPS 只支持这一种
an.AnimationOrder = n     # 同一个 n 的形状合并成同一级，一起出现
an.AdvanceMode = 2        # 2 = 上一级播完自动播下一级（XML 里是 nodeType="afterEffect"）
an.AdvanceTime = 0.12     # 每级之间的间隔，秒
```

要点：

- `AdvanceMode = 2` 是做出「自动依次展开」的关键；不设就是 `clickEffect`，每一级都要手动点。
- 同一 `AnimationOrder` 的多个形状会合并成一级。所以**表格按行、卡片按张**分配级号，
  不要给每个形状单独一级，否则形状多一点的页面会拖出几十级。
- 实测放映：内容正常显示，翻页正常，动画不会卡住翻页。

### 交付前校验（保存在 python 侧看不出效果，直接查 XML）

```python
import re, zipfile
z = zipfile.ZipFile("out.pptx")
pat = r"<p:transition[^>]*>(?:<p:([a-zA-Z]+))?"
for n in sorted(x for x in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", x)):
    xml = z.read(n).decode("utf-8")
    m = re.search(pat, xml)
    print(n, m.group(1) if m else None,
          "click=%d after=%d" % (xml.count('nodeType="clickEffect"'),
                                 xml.count('nodeType="afterEffect"')))
```

### 视觉自检：用 WPS 自己出图

本机没有 LibreOffice，`artifact-preview` 对 pptx 出不了图（对 pdf 还缺 PyMuPDF）。
改用 WPS 自己导出每页 PNG：

```python
pres.Slides(i).Export(r"out\p01.png", "PNG", 1280, 720)
```

---
## CLI 命令结构

```
wps
├── document          # 文档管理: new/open/save/info
├── writer            # 文字: add-paragraph/heading/list/table/image
├── calc              # 表格: set-cell/get-cell/set-range/merge-cells
├── impress           # 演示: add-slide/remove-slide/set-content/add-element
├── style             # 样式: create/modify/list/apply/remove
├── preset            # 设计预设: list/info/apply
├── export            # 导出: render output.pptx -p pptx
└── session           # 会话: status/undo/redo/history
```

### 使用示例

```bash
# 创建文档
cli-anything-wps document new --type writer --name "报告" -o report.json
cli-anything-wps --project report.json writer add-heading -t "年度报告" -l 1
cli-anything-wps --project report.json export render report.docx -p docx

# JSON模式（Agent使用）
cli-anything-wps --json document new --type writer --name "test"
```

### 设计预设

| 预设 | 配色 | 适用场景 |
|------|------|---------|
| academic | 深蓝#1A3C8B | 学术答辩/基金申请 |
| consultant | 深蓝#003366+亮青 | 咨询报告 |
| business | 商务蓝#005294 | 会议汇报/课件 |
| tech | 近黑#0F1423 | 科技/AI/数据报告 |

### 质量审查（5维度）

| 维度 | 检查项 | 阈值 |
|------|--------|------|
| visual | 字体层级/对比度/留白 | 70分 |
| pedagogy | 叙事弧/预备知识 | 75分 |
| proofreading | 拼写/术语/溢出 | 80分 |
| parity | PPTX/PDF一致性 | 85分 |
| substance | 数据准确性/引用 | 90分 |
