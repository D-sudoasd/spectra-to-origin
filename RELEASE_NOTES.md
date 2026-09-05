# SpectraToOrigin v1.0.0

Windows 双击即可使用，不必先安装 Python。

## 下载

资源里的 **SpectraToOrigin.exe**（约 18 MB）。放到任意文件夹，双击打开。

1. 把谱线 txt 或整个文件夹拖进窗口
2. 确认自动识别的 XYYY / XYXY
3. 点「生成 Origin 工程 (.opju)」

## 说明

- **不需要 Python**
- 写出 `.opju` 需要本机已安装 **Origin Pro**（经 Origin 2025 验证）
- 没有 Origin 时仍可加载谱线，并勾选「同时导出 xlsx/csv」
- 命令行同样可用：`SpectraToOrigin.exe --cli --xlsx-only -i 谱线文件夹 -o out.xlsx`

请从本 Release 下载 exe，不要用源码树里的 `dist/`（未纳入 git）。
