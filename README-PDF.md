# PDF 预览功能使用说明

## 使用方法

直接双击打开 `index.html` 文件即可使用，无需启动 HTTP 服务器。

## 功能说明

PDF viewer 使用 PDFium.js 提供以下功能：

- ✅ 预览 PDF 文件（高清晰度显示，基于 Google PDFium 引擎）
- ✅ 页面导航（上一页、下一页、跳转）
- ✅ 缩放控制（放大、缩小、适应页面）
- ✅ 下载功能
- ✅ 打印功能
- ✅ 高质量渲染（使用 WebAssembly 实现）

## 实现方式

使用 PDFium.js（@hyzyla/pdfium）：
- 使用 CDN 版本的 PDFium.js 库
- 基于 Google PDFium 引擎的 WebAssembly 实现
- 支持高清晰度 PDF 渲染
- 支持直接打开 HTML 文件进行本地预览

## 文件结构

主要文件：

```
office/
├── index.html
├── pdf-viewer.js
└── ...
```

## 故障排除

如果 PDF 无法加载，请检查：

1. ✅ 浏览器控制台是否有错误信息
2. ✅ 网络连接是否正常（需要加载 CDN 资源）
3. ✅ 浏览器是否支持 WebAssembly（现代浏览器都支持）
4. ✅ 浏览器是否支持 ES modules

## 参考

- PDFium.js 文档：https://pdfium.js.org/
- @hyzyla/pdfium：https://www.npmjs.com/package/@hyzyla/pdfium
