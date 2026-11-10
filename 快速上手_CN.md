# 快速上手

1. 使用已有的 Conda `test` 环境（Python 3.12 x64）；运行脚本会激活它，不会新建环境。
2. 解压源码包。
3. 双击 `run_windows.bat` 启动程序（自动激活 `test` 环境并安装所需依赖）。
4. 默认浏览器会打开本机 WebUI，将 PPTX、PDF 或 DOCX 拖到页面预览区。
5. 左侧根据页面缩略图切页；右侧复选框列出整份文档的图片和表格，勾选状态会跨页保留。
6. 文字不在选择范围中。点击页面里的图片或表格框也可将对象加入勾选。
7. 按 `Ctrl+C` 或点击“复制勾选项”，再到目标软件按 `Ctrl+V`。

## 生成 EXE

在 Windows 上双击 `build_windows.bat`，使用 `test` 环境生成文件：

`dist\文档对象工坊_全篇勾选版.exe`

## 注意

- PPTX 组合/形状原生复制需要 Microsoft PowerPoint。
- Word 原生表格复制需要 Microsoft Word；PPTX 表格/图表原生复制需要 Microsoft PowerPoint。
- LibreOffice 可作为预览后备。
- 预览按 3 倍比例渲染；打开时会显示上传、解析和渲染进度。
- 多选内容会作为 HTML 集合复制；单选图片保留原图数据。
- EXE 会启动本机 WebUI 并打开默认浏览器；实际跨软件剪贴板粘贴仍需验收。
