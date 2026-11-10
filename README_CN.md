# 文档对象工坊（Document Object Studio）v0.2.0

Windows 本地 WebUI，用于从 PPTX、PDF、DOCX 中勾选整份文档里的图片与表格，再复制到 **Windows 系统剪贴板**。文字对象不在选择列表中，也不能从页面点击选择。浏览器只负责界面，文档解析、Office COM 和剪贴板操作都在本机 Python 后端完成。

> **当前交付状态：WebUI 源码、Windows 构建脚本和 EXE。** EXE 会启动本机服务并打开默认浏览器。本机三份样本已完成解析、全篇图片/表格勾选、跨页保持、复制路由和浏览器渲染检查；Office 内实际粘贴仍需验收。

## 1. 支持格式

- `.pptx`：PowerPoint 演示文稿。
- `.pdf`：可搜索 PDF、含图片 PDF、含矢量路径 PDF。
- `.docx`：Word 文档（不支持旧版 `.doc`）。

页面预览以 3 倍比例渲染，左侧显示对应页面的真实缩略图。对象框使用文档原始坐标和页面坐标范围进行归一化映射，缩放时不会偏移；页面上只突出显示悬停或选中的对象。打开过程显示上传、解析和渲染进度。

右侧列表汇总整份文档的图片与表格，每项带有复选框和所在页码。勾选状态跨页面保留；可从右侧清单或页面中的图片/表格边界框进行选择。文字、普通形状和公式不会进入可复制对象列表。

## 2. 当前实现

### PPTX
- 通过 PowerPoint COM（Windows）或 LibreOffice（后备/测试环境）导出页面预览。
- 递归解析幻灯片对象树；右侧只提供图片、表格和图表供复制。
- PowerPoint 原生表格和图表按完整对象复制；图片从 PPTX 内部媒体直接提取，不截图、不按幻灯片预览分辨率缩放。
- 图片复制写入原始图像 MIME 数据；原始图片像素不由本程序重采样。

### PDF
- 多页预览和页面导航。
- 文字按行解析用于页面识别，但不可勾选或复制。
- 图片按 PDF 图像资源和页面放置位置提取，尽可能保留 PDF 内嵌原始图像字节。
- 表格由 PyMuPDF 按固定的线条检测路径识别，复制为 HTML 表格和纯文本。
- 宽度或高度小于 2.5 点的图片片段不进入清单，避免将 PDF 中拆成细条的文字图像当成插图。

### DOCX
- 使用 Word COM（Windows）或 LibreOffice（后备/测试环境）生成分页预览。
- 解析正文段落、完整表格、内嵌图片和 OLE 对象/公式；右侧只允许勾选图片和表格。
- 跨页表格只显示一个完整勾选项，分页预览中的表格框会映射回同一原生表格。
- 页面预览对象来自 Word 导出的 PDF；原始图片和 OLE 公式预览可映射回 DOCX 包内媒体资源。
- 图片复制使用 DOCX 包内原始图片数据。
- OLE 公式在 Windows + Word 下调用 Word 原生对象复制；无 Word 时尝试复制 WMF 矢量预览（需 LibreOffice + Inkscape；PNG 回退还需 CairoSVG）。
- Windows + Word 下，单选表格时调用 Word 原生 Range.Copy，保留完整表格结构。

## 3. 快速启动（Windows）

### 环境
- Windows 10/11 x64。
- 已安装 Conda，并已有名为 `test` 的 Python 3.12 x64 环境。
- Microsoft Office 桌面版（推荐；PPTX/Word 原生对象复制依赖它）。
- LibreOffice 可作为 Office COM 不可用时的预览后备；WMF 矢量预览转换还建议安装 Inkscape。

### 运行源码
1. 解压本压缩包。
2. 双击 `run_windows.bat`；脚本会激活已有的 Conda `test` 环境，并在其中安装项目依赖。
3. 默认浏览器会打开本机 WebUI；将 `.pptx`、`.pdf` 或 `.docx` 拖入预览区。

### 构建 EXE
1. 在 Windows 上双击 `build_windows.bat`；脚本会使用 Conda `test` 环境。
2. 等待依赖安装和 PyInstaller 打包。
3. EXE 输出到 `dist\文档对象工坊_全篇勾选版.exe`。

构建过程不会创建独立 `.venv`，依赖会安装到已有的 `test` 环境。

## 4. 使用方式

1. 打开程序，点击“打开文档”或拖入文档。
2. 左侧根据页面缩略图选择页码；左侧方向键上/下或左/右均可翻页。
3. 在右侧全篇列表勾选图片或表格；也可点击页面中的图片/表格边界框。文字不会被选中。
4. 勾选状态跨页面保留。单选图片保留原图数据；多选内容按文档顺序以 HTML 合并复制。
5. 点击“复制勾选项”或按 `Ctrl+C`，再到目标软件按 `Ctrl+V`。

## 5. 剪贴板格式说明

| 对象 | 复制格式 |
|---|---|
| 单张图片 | 原始图片 MIME + Windows DIB 兼容格式 |
| 多张图片/多张表或混合选择 | 保留原图数据并以 HTML 合并，附纯文本回退 |
| PDF 表格 | `text/html` + `text/plain` |
| PPTX 表格/图表 | Windows PowerPoint 原生对象复制 |
| DOCX 表格 | Windows Word 原生整表复制；兼容路径为 HTML/纯文本 |

WebUI 后端通过 Windows 剪贴板格式发布文本、HTML、图片和 SVG，并通过 Office COM 发布 PowerPoint/Word 原生对象。目标软件会自行决定采用哪个格式。

## 6. 测试

- 详细功能边界：`功能说明_CN.md`
- 快速操作：`快速上手_CN.md`
- 本轮回归测试报告：`测试报告_CN.md`
- 测试脚本：`tests\test_sample_documents.py` 和 `tests\test_webui_api.py`
- Windows 测试入口：`run_tests_windows.bat`

将 PDF、PPTX、DOCX 样本放入 `test_samples` 文件夹，然后运行 `run_tests_windows.bat`。也可以设置 `DOCSTUDIO_SAMPLE_DIR` 指向样本所在文件夹。测试会检查解析、高清预览、页面缩略图和 WebUI 复制路由，并在本地生成被 Git 忽略的 JSON 检查结果。

## 7. 当前限制

- 这不是 PowerPoint、Acrobat 或 Word 的完整替代品。
- PDF 表格识别依赖文本和线条布局；扫描表格、复杂合并单元格或无边框布局可能无法准确识别。
- PDF 矢量路径不进入图片/表格勾选清单；当前清单列出内嵌图片和自动识别表格。
- 扫描 PDF 若整页只有一张扫描图，无法从中自动恢复原始文字/独立图片对象；需要 OCR 或人工框选。
- DOCX 页内对象位置来自 Word/LibreOffice 的分页 PDF 预览。DOCX 原始媒体与预览图的自动映射仅在对象数量一致时按文档顺序映射；其他文档如果对象重复、跨页、页眉页脚、浮动锚点或导出器重排，映射可能回退到 PDF 渲染资源。
- DOCX 复杂浮动形状、SmartArt、嵌套组合、OLE 激活编辑和页眉页脚对象仍需扩展；本流程仅提供图片与表格。
- PPTX SmartArt、图表内部数据、OLE、3D、动画、复杂裁剪及深层嵌套组合需在目标 Office 版本上逐类验收。
- 本机已用 Word/PowerPoint COM 生成三份样本的页面预览并构建 EXE；剪贴板格式在其他软件中的实际粘贴效果仍需手工验收。

## 8. 文件结构

```text
DocObjectStudio_Source/
  web/                    HTML/CSS/JavaScript 前端
  document_engine.py      PDF/PPTX/DOCX 解析、预览与 payload 核心
  webui.py                本机 WebUI 服务和 Windows 剪贴板桥接
  requirements.txt        Python 依赖
  run_windows.bat         源码运行入口
  build_windows.bat       Windows EXE 构建入口
  run_tests_windows.bat   样本回归测试入口
  功能说明_CN.md          功能与操作说明
  测试报告_CN.md          本轮测试结论
  tests/test_sample_documents.py 三份样本解析和高清预览检查
  tests/test_webui_api.py 本地 WebUI 接口检查
```
