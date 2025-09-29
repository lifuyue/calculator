# glycoenum CLI

`glycoenum` 枚举给定寡糖组成的所有唯一排列，输出糖链分子式、终端修饰后的分子式以及理论质量。程序固定内置六个单元（Hex、deoxyhex、pent、HexN、UA、HexNAc）以及一次性的 2 PMP 终端修饰（C20H18N4O）。

## 快速开始

安装依赖并调用：

```bash
python -m pip install --upgrade pip
python -m pip install -e .
```

两种计数写法互通，如下等价：

```bash
# 命名参数
python -m glycoenum.cli --hex 5 --deoxyhex 0 --pent 1 --hexn 0 --ua 0 --hexnac 0 \
  --adduct "[M+H]+" --csv out.csv

# 位置参数
python -m glycoenum.cli 5 0 1 0 0 0 --adduct "[M+H]+"
```

计算完成后，终端会询问是否在当前目录生成 `glycoenum_output.xlsx`；回答 `Y` 即可得到与 CSV 同列的表格。

所有行都会采用表头 `compound,分子式,最终分子式,理论`，并在 `--csv` 缺省时输出到标准输出。`--max-rows` 可用于提前截断（超过时在 stderr 打印 `[warn] rows truncated at …`）。

## 质量模型与修饰

- `--mass-model`：`monoisotopic`（默认）或 `average`。
- `--masses`：覆写原子质量，例如 `--masses "C=12.0,H=1.007825,N=14.003074,O=15.994915"`。
- `--adduct`：`neutral`、`[M+H]+`、`[M+Na]+`。
- `--decimals`：控制理论质量的小数位，默认 4。

`n = Hex + deoxyhex + … + HexNAc` 需满足 `2 ≤ n ≤ 10`。分子式遵循 `Σ(单元)` → 扣除 `(n−1)·H₂O` → 加一次 2 PMP。

## 打包

使用 PyInstaller 生成单文件可执行程序：

```bash
pyinstaller --noconfirm --onefile glycoenum/cli.py -n glycoenum
```

Windows 环境可直接运行 `build.bat`，脚本将检测 Python、安装必须依赖并在 `dist\glycoenum.exe` 输出结果。
