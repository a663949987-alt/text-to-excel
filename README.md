# 文本转 Excel

一个小巧的 Windows 桌面工具：把 TXT、CSV、TSV 文件或粘贴文本按字段拆分，并导出为样式清晰的 Excel 工作簿。

## 功能

- 将 `.txt`、`.csv`、`.tsv` 文件直接拖入窗口
- 支持粘贴文本与即时表格预览
- 自动识别 Tab、英文/中文逗号、竖线、分号和连续空格
- 支持 UTF-8、GB18030/GBK、UTF-16 文件编码
- 首行可作为表头，可忽略空行并清理字段前后空格
- Excel 自动添加深色表头、筛选、冻结首行、边框、交替行色和合适列宽
- 全程本地处理，不上传文本内容

## 使用

在 [Releases](https://github.com/a663949987-alt/text-to-excel/releases/latest) 下载：

- `Excel-win-x64.exe`：直接运行
- `Excel-win-x64.zip`：下载更快，解压后运行其中的 `文本转Excel.exe`

程序无需安装，适用于 64 位 Windows。

## 从源码构建

需要 .NET 10 SDK：

```powershell
dotnet restore
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

## 开源许可

MIT License
