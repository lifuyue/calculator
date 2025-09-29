# 模块与接口说明文档

## 项目概览
- **名称**：glycoenum
- **目标**：枚举顺序敏感的寡糖序列，计算分子式、修饰后分子式与理论质量，并导出 CSV。
- **运行形态**：Typer 驱动的命令行工具，可后续封装为单文件可执行程序。
- **固定输出字段**：`compound`、`分子式`、`最终分子式`、`理论`，其中分子式遵循 Hill 顺序（C、H 置前，其余按字母序）。

## 公共类型（`glycoenum/types.py`）
- `UnitDefinition(name: str, formula: str, metadata: Mapping[str, Any] = {})`
- `Modifier(label: str, formula: str)`
- `Composition(atoms: Mapping[str, int])`
- `SequenceEntry(sequence: tuple[str, ...], composition: Composition, final_composition: Composition, mass: float, metadata: dict[str, Any] = {})`
- `MassModel(name: str, atom_masses: Mapping[str, float])`
- `RuntimeConfig` 字段：`units`、`modifier`、`nmin`、`nmax`、`mode`、`adduct`、`mass_model_name`、`mass_overrides`、`csv_path`、`filters`、`max_rows`、`verbose`、`locale`

## 模块职责与接口

### `glycoenum/formula_parser.py`
- **职责**：解析任意分子式，支持多式合并、桥连失水、修饰叠加，并按 Hill 顺序输出。
- **接口**：
  - `parse_formula(formula: str) -> Composition`
  - `sum(formulas: Sequence[str]) -> dict[str, int]`
  - `dehydrate(counts: Mapping[str, int] | Composition, n: int) -> dict[str, int]`
  - `add_modifier(counts, modifier: str) -> dict[str, int]`
  - `format_hill(counts) -> str`

### `glycoenum/mass_calc.py`
- **职责**：维护质量模型，支持原子质量覆写，解析加合物表达式。
- **接口**：
  - `resolve_mass_model(name: str, overrides: Mapping[str, float] | None = None, *, default_models_path: Path | None = None) -> MassModel`
  - `composition_mass(composition: Composition, model: MassModel) -> float`
  - `apply_adduct(base_mass: float, adduct: str | None, model: MassModel) -> float`

### `glycoenum/enumerator.py`
- **职责**：根据单元、修饰和长度范围枚举顺序敏感序列，处理桥连失水、修饰加成，并触发质量计算。
- **核心类**：`SequenceEnumerator`
  - 构造：`SequenceEnumerator(units, modifier, base_mass_model, *, adduct=None, filters=())`
  - 方法：`generate(nmin, nmax, *, mode="sequences", max_rows=None) -> Iterator[SequenceEntry]`

### `glycoenum/config.py`
- **职责**：解析 CLI 参数与外部配置，载入默认单元、修饰，并生成 `RuntimeConfig`。
- **接口**：
  - `build_runtime_config(cli_args: Namespace | Mapping[str, Any]) -> RuntimeConfig`
  - `load_units(path: Path | None, inline_data: str | None = None) -> list[UnitDefinition]`
  - `parse_modifier(text: str | None) -> Modifier | None`
  - （内部）`_parse_units_text`、`_parse_units_path`、`_parse_masses`、`_collect_filters`

### `glycoenum/cli.py`
- **职责**：定义 Typer 应用，串联配置层、枚举器与 CSV 输出。
- **接口**：
  - `run(...)` — 主命令，解析参数、加载配置、执行枚举与导出。
  - `main()` — Typer 入口（`pyproject.toml` 中暴露为 `glycoenum`）。

### 其它支持
- `glycoenum/defaults/units.toml`：提供 6 个默认单元（pent、hex、hexn、dhex、neuac、neugc）。
- 默认修饰：`Modifier(label="2 PMP", formula="C18H22N2O6")`，可通过 CLI 覆盖。

## 模块交互
```
CLI (Typer)
  └── config.build_runtime_config → RuntimeConfig
        ├── load_units / parse_modifier
        └── mass override 解析
  └── mass_calc.resolve_mass_model → MassModel
  └── SequenceEnumerator.generate → SequenceEntry 流
        ├── formula_parser.parse_formula 获取单元/修饰组成
        ├── composition_mass + apply_adduct 计算质量
        └── 过滤/行数控制
  └── CSV 输出 (compound, 分子式, 最终分子式, 理论)
```

## CLI 参数表
| 参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `--units` | `str \| Path` | `defaults/units.toml` | 单元定义文件路径或内联 JSON/TOML；可完全覆写默认 6 单元。 |
| `--modifier` | `str` | `2 PMP` | 修饰分子式或 `label=formula` 形式，留空表示不加修饰。 |
| `--nmin` | `int` | `2` | 序列最小长度。 |
| `--nmax` | `int` | `10` | 序列最大长度。 |
| `--mode` | `str` | `sequences` | 枚举模式（当前仅支持顺序敏感序列）。 |
| `--adduct` | `str` | `[M+H]+` | 加合物表达式，留空输出中性质量。 |
| `--mass-model` | `str` | `monoisotopic` | 质量模型名称（内置 `monoisotopic`/`average`，可扩展）。 |
| `--masses` | `str` | 空 | 原子质量覆写列表，如 `C=12.0,H=1.007825`。 |
| `--csv` | `Path` | 空（stdout） | CSV 输出路径；缺省输出到标准输出。 |
| `--filter` | `str`（可多次） | 空 | 序列名称子串过滤，支持多次传入。 |
| `--max-rows` | `int` | 空 | 限制输出行数，控制枚举规模。 |
| `--locale` | `str` | `zh-CN` | 输出编码提示（`zh-*` 时采用 `utf-8-sig`）。 |
| `--config` | `Path` | 空 | JSON/TOML 综合配置文件，CLI 参数可覆写其中字段。 |
| `--verbose` | `bool` | `False` | 打印质量模型、过滤器等诊断信息。 |
| `--version` | flag | - | 显示版本并退出。 |

## 计算约定
- `分子式` = Σ(单元分子式) − (n − 1) × `H₂O`。单元序列顺序敏感，例如 `pent-hex-hex-hex-hex`。
- `最终分子式` = `分子式` + 修饰分子式（整体仅叠加一次）。
- 质量计算基于 `最终分子式`，再应用 `--adduct`。

## 扩展建议
- 后续支持更多枚举模式（如组合去重复、重复次数限制）。
- 引入外部质量模型文件加载和国际化输出。
- 编写单元测试确保公式解析、加合物处理和过滤行为的正确性。
