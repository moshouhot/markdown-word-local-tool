# Markdown → Word 本地克隆

这是一个本地单页版 Markdown → Word 转换工具。

## 使用方式

- 直接双击 `index.html`，或
- 在本目录运行：

```powershell
python -m http.server 8000
```

然后打开 `http://127.0.0.1:8000/`。

## 本地化说明

- 原页面的核心样式与脚本均在 `index.html` 内联，Markdown 转 Word / docx 重排逻辑可离线运行。
- 已移除远程百度统计脚本，避免本地打开时发起统计请求。
- 已删除页脚推广和其他外部跳转入口。
- 保留转换相关功能，删除品牌页脚、推广和外部跳转入口。


