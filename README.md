# Trinkter —— 超声设备担保申报 PDF 自动填写工具

## 简介

Trinkter 是一个基于 Python + Tkinter 的桌面工具，用于自动填写德国 KV（Kassenärztliche Vereinigung）超声设备担保申报 PDF 表单。它会根据用户输入的设备型号和探头型号，自动勾选对应的 AK（Anwendungsklasse）checkbox，并填写基本信息字段。

---

## 文件结构

```
Gwe/
├── Trinkter_new.py      # 主界面（Tkinter GUI）
├── Pdf_filling.py       # PDF 读写核心逻辑
├── Excel.py             # 探头 → AK 映射（读取 Excel）
├── data.xlsx            # 探头 → AK 对照表（Excel.py 读取）
├── checkcheck.py        # 独立工具：检测 PDF 模板中 AK 表单字段是否齐全
├── requirements.txt     # Python 依赖列表
└── README.md            # 本文档
```

---

## 使用方法

### 1. 运行程序

```bash
python Trinkter_new.py
```

### 2. 界面操作步骤

| 步骤 | 字段 | 说明 |
|------|------|------|
| ① | PDF文件 | 点击"浏览"选择要填写的 KV 担保申报 PDF |
| ② | 公司名称 | 填写 Adresse（公司/机构名称） |
| ③ | 设备型号 | 从下拉菜单选择设备，设备名称和探头列表会自动更新 |
| ④ | 设备名称 | 自动填入（选择设备型号后自动填写，可手动修改） |
| ⑤ | 设备号(SN) | 输入序列号，例如 `360101-M26305870006` |
| ⑥ | 生产年份 | **自动**从 SN 号解析（只读，无需手动填写） |
| ⑦ | 日期 | **自动**填入当天日期（格式 DD.MM.YYYY，可修改） |
| ⑧ | 探头选择 | 在左侧列表选中探头，点击"→"移到右侧；右侧为实际使用的探头 |
| ⑨ | 输出文件 | 填写输出 PDF 的路径/文件名 |
| ⑩ | 开始 | 点击"开始"执行填写 |

---

## SN 号解析规则

SN 格式示例：`360101-M26305870006`

程序会提取 `M` 后面的前两位数字作为年份后两位：

- `M26…` → 生产年份 `2026`
- `M23…` → 生产年份 `2023`

如果 SN 格式不符合规则（没有 `M` + 两位数字），生产年份字段留空。

---

## 如何添加新设备及探头映射

在 `Trinkter_new.py` 的 `__init__` 方法中找到 `self.device_map` 字典，按如下格式添加新设备：

```python
self.device_map = {
    # 已有设备 ...

    "新设备型号名称": {
        "name": "PDF中显示的完整设备名称",
        "probes": ["探头1", "探头2", "探头3"]
    },
}
```

**示例：**

```python
"Acclarix AX8 Series": {
    "name": "Acclarix AX8 Diagnostic Ultrasound System",
    "probes": ["C5-2Q", "L17-7HQ", "E8-4Q", "P5-1Q"]
},
```

添加后，新设备会自动出现在下拉菜单中，选择后探头列表自动更新。

---

## 探头频率映射

在 `Pdf_filling.py` 中有一个 `frequency_map` 字典，用于将探头型号映射到频率字符串：

```python
frequency_map = {
    "L743-2": "4-3 MHz",
    "C361-2": "6-3 MHz",
    # 添加新探头 →
    "新探头型号": "X-Y MHz",
}
```

如果探头型号符合 `数字-数字` 格式（如 `L17-7HQ`），程序会自动解析频率（`17-7 MHz`），无需手动添加。只有格式特殊的探头才需要在 `frequency_map` 中手动指定。

---

## PDF 字段说明

`Pdf_filling.py` 的 `fill_pdf_fields` 函数填写以下字段：

| PDF字段名 | 内容 | 备注 |
|-----------|------|------|
| `Adresse` / `Adresse_1` | 公司名称 | 两页均填 |
| `Bezeichnung` / `Bezeichnung_1` | 设备名称 | 两页均填 |
| `GeräteNummer` / `GeräteNummer_1` | 设备号(SN) | 两页均填 |
| `Baujahr` / `Baujahr_1` | 生产年份 | 两页均填 |
| `Auslieferungsdatum` / `Auslieferungsdatum_1` | 交货日期 | 两页均填 |
| `datum` | 当前日期 | 表单日期签名处 |
| `Schallkopf 1`~`Schallkopf 5` | 探头型号 | 最多5个 |
| `Frequenz 1`~`Frequenz 5` | 探头频率 | 与探头对应 |
| `ak_X_Y` | AK checkbox | 由 Excel 查询结果决定 |

若需新增填写字段，在 `fill_pdf_fields` 函数中仿照以下格式添加：

```python
set_text_field(pdf, "新字段名", 变量值)
```

---

## AK checkbox 勾选逻辑

1. 用户在界面选择探头后，`Excel.find_items()` 根据探头列表查询 Excel 表，返回应勾选的 AK 编号列表（如 `["AK 2.1", "AK 3.4", ...]`）
2. 程序将 AK 编号标准化为 `ak_X_Y` 格式（`normalize_name` 函数）
3. 遍历 PDF 所有 checkbox 字段，名称匹配则勾选

**特殊规则：**
- 设备名称为 `DUS 60` 时，自动追加 `AK 20.12`（Farbkodierte nein）
- 其他设备自动追加 `AK 20.11`（Farbkodierte ja）

---

## checkcheck.py —— PDF 模板字段检测工具（独立使用）

`checkcheck.py` **不属于主程序**（`Trinkter_new.py` 不会调用它），是一个独立的命令行小工具，用于在使用新的 KV PDF 模板之前，检测该 PDF 中是否包含所有需要的 AK checkbox 表单字段（`ak_1_1`、`ak_2_3` 这类字段名）。

**工作原理：**

1. 用 PyPDF2 读取 PDF 中所有表单字段名
2. 生成完整的 AK 编号清单（AK 1.1 ~ AK 23.1，共约 50 项）
3. 将每个 AK 编号标准化为 `ak_X_Y` 格式，逐一检查 PDF 中是否存在同名字段
4. 输出每项的 ✓/✗ 结果和总完成度百分比

**使用方法：**

打开 `checkcheck.py`，把开头的 `PDF_PATH` 改成要检测的 PDF 文件名：

```python
PDF_PATH = "KV Hessen.pdf"
```

然后运行：

```bash
python checkcheck.py
```

输出示例：

```
AK 1.1   ✓
AK 2.1   ✓
AK 20.11 ✗
...
完成度: 48/50 (96.0%)
```

带 ✗ 的项说明 PDF 模板中缺少对应的 checkbox 字段，主程序填写该模板时这些 AK 将无法被勾选，需要先在 PDF 编辑器中补齐字段（字段名格式为 `ak_X_Y`）。

---

## 依赖安装

```bash
pip install -r requirements.txt
```

包含：`pdfrw`（主程序 PDF 读写）、`PyPDF2`（checkcheck.py）、`pandas` + `openpyxl`（Excel.py 读取 data.xlsx）。

（其余为标准库：`tkinter`、`re`、`threading`、`datetime`）

---

## 常见问题

**Q: 为什么某些 checkbox 没有被勾选？**

A: 检查 PDF 中该字段的实际名称是否与 `ak_X_Y` 格式一致。可在 `Pdf_filling.py` 中查看控制台打印的"目标字段"和匹配结果。

**Q: 如何更换 PDF 模板？**

A: 直接在界面选择新的 PDF 文件即可。不同 KV 的 PDF 字段名需与 `fill_pdf_fields` 中的字段名一致。如有差异，需在 `Pdf_filling.py` 中相应修改字段名。

**Q: 日期格式可以改吗？**

A: 可以。在 `Trinkter_new.py` 中找到：

```python
self.datum_var = tk.StringVar(value=date.today().strftime("%d.%m.%Y"))
```

将 `"%d.%m.%Y"` 改为所需格式，如 `"%Y-%m-%d"` 即为 `2026-06-29` 格式。
