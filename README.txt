谱线 → Origin 工程 (.opju)
=======================

把多条两列 1D 谱直接写进 Origin Pro：同一张工作表、XYYY 或 XYXY 列类型，并生成曲线图，保存为 .opju。

需要
----
本机已安装 Origin Pro（当前机器是 Origin 2025），且当前 Python 能 import originpro。

启动
----
双击 run.bat（推荐）或 启动.bat
或：py -3 spectra_to_origin.py

不要用带中文的 chcp 65001 批处理；脚本目录在「桌面」下时，旧启动文件会把命令拆断。

用法
----
1. 把谱线 txt 或文件夹拖进窗口（也可点添加文件/文件夹）
2. X 列相同会自动识别为 XYYY；不同则为 XYXY。仍可手动改
3. 加载后按文件名自动分组（不限于温度）。也可均分成 3/4 张，或把选中项移入某组、改组名
4. 点「生成 Origin 工程 (.opju)」：每组一张工作表 + 一张线图

同一文件夹里的「按温度」子目录是副本，添加上级文件夹时会跳过，避免导两遍。

命令行
----
py -3 spectra_to_origin.py --cli -i "谱线文件夹" -o "out.opju" --layout XYYY --split-temp

这批 Ti-2448 合成谱 X 网格相同，用 XYYY。不要从峰面积读相分数，不要从峰宽读晶粒。
