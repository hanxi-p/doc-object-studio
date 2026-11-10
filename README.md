# 文档对象工坊（Document Object Studio）

Windows 本地文档对象工具。用浏览器打开 PPTX、PDF、DOCX，在整份文档的清单中勾选图片、图表或表格，并复制到剪贴板。

文档在本机解析和预览，不需要上传到在线服务。当前版本提供 WebUI、Windows EXE 和完整源码。

## 功能

- 3 倍高清页面预览、左侧页面缩略图和加载进度。
- 右侧按整份文档列出图片与表格，并以复选框跨页多选。
- 页面上的文字不属于选择范围；只允许勾选图片、图表和表格。
- 保留嵌入图片原始字节；支持 HTML 多选复制及 Office 原生表格/图表复制。
- 预览边界框按文档坐标定位，支持键盘翻页和分栏滚动。

## 下载

前往 [GitHub Releases](https://github.com/hanxi-p/doc-object-studio/releases/latest) 下载 Windows 版 EXE 和源码压缩包。

EXE 会启动本机 WebUI 并打开默认浏览器。使用 PowerPoint 或 Word 原生复制功能需要安装 Microsoft Office。

## 从源码运行

Windows 10/11，Python 3.12。此项目构建脚本使用 Conda `test` 环境：

1. 安装依赖：运行 `run_windows.bat`。
2. 构建 EXE：运行 `build_windows.bat`。
3. EXE 输出到 `dist/文档对象工坊_全篇勾选版.exe`。

更详细的中文说明见 [README_CN.md](README_CN.md)、[快速上手_CN.md](快速上手_CN.md) 和 [功能说明_CN.md](功能说明_CN.md)。

## 测试

用户文档不会随仓库提供。将包含图片和表格的 PDF、PPTX、DOCX 测试文件放入 `test_samples`，再运行 `run_tests_windows.bat`。也可以通过 `DOCSTUDIO_SAMPLE_DIR` 指定本地样本目录。

## 当前限制

- PDF 表格识别依赖线条与文本布局；扫描件或无边框表格可能需要手动处理。
- Office 原生复制依赖本机 Microsoft Office。
- 源码未附带开源许可证。
