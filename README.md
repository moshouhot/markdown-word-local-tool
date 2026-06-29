# Markdown → Word 本地工具

这是一个本地运行的 Markdown → Word 转换工具。核心功能保留为前端 HTML/CSS/JS，Windows 桌面版通过 C# WebView2 壳加载同一套页面。

## 运行网页版

直接打开 `index.html`，或在本目录运行：

```powershell
python -m http.server 8000
```

然后打开 `http://127.0.0.1:8000/`。

## 运行 Windows EXE

已生成的本地发布包：

- 发布目录：`dist/MarkdownWordTool-win-x64/`
- 压缩包：`dist/MarkdownWordTool-win-x64.zip`
- 启动程序：`dist/MarkdownWordTool-win-x64/MarkdownWordTool.exe`

> Windows 10/11 通常已包含 Microsoft Edge WebView2 Runtime；如果目标电脑缺失，需要安装 WebView2 Runtime。

## 开发结构

```text
index.html                 页面结构入口
assets/styles.css          样式
assets/app.js              Markdown 解析、DOCX 构建、界面交互
windows-shell/             C# WinForms + WebView2 Windows 壳
_analysis/                 验证脚本和验证结果
```

WebView2 壳使用固定虚拟来源加载前端：

```text
https://app.local/index.html
```

这样 EXE 内的 `localStorage` 来源稳定，不受文件路径或本地 HTTP 端口变化影响。

## 设置存储

网页版设置存储在当前浏览器当前来源的 `localStorage`。

Windows EXE 设置存储在 WebView2 用户数据目录：

```text
%APPDATA%\MarkdownWordTool\WebView2
```

参数区提供“导出设置 / 导入设置”，可把自定义模板、主题、上次模板、侧栏宽度和提示偏好备份为 JSON；不会保存正文内容或上传文件。

## 构建 Windows EXE

```powershell
dotnet build .\windows-shell\MarkdownWordTool.csproj -c Release
```

输出：

```text
windows-shell/bin/Release/net8.0-windows/MarkdownWordTool.exe
```

## 生成可分发包

```powershell
dotnet publish .\windows-shell\MarkdownWordTool.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=false `
  -o .\dist\MarkdownWordTool-win-x64

Compress-Archive -Path .\dist\MarkdownWordTool-win-x64\* `
  -DestinationPath .\dist\MarkdownWordTool-win-x64.zip `
  -CompressionLevel Optimal
```

## 验证

```powershell
node --check .\assets\app.js
node .\_analysis\docx-build-smoke.cjs
dotnet build .\windows-shell\MarkdownWordTool.csproj -c Release
```

EXE 交互验证脚本：

```text
_analysis/cdp-verify.cjs
_analysis/webview2-verify-result.json
```

## 本地化说明

- Markdown 转 Word / docx 重排逻辑可离线运行。
- 已移除远程统计、页脚推广和其他外部跳转入口。
- 页面聚焦本地文档处理，内容不会上传服务器。

## 交互增强

- 输入区新增“插入示例”，可快速试用 Markdown 转 Word 流程。
- 模式切换区新增当前模式说明，降低模板 / 净化 / 重排的理解成本。
- 标题 H1/H2/H3 细项默认折叠到“高级标题设置”，常用参数更清爽。
- 右侧转换、清空、PDF 和状态提示固定为底部操作区，减少滚动寻找按钮。
- 清空内容支持 10 秒内撤销，避免误删长文本。
- 参数区新增“导出设置 / 导入设置”，方便长期迁移设置。
