# single-table-converter

把**一份**供应商 SKU 报价单 / 库存表,交给 AI 一次性转成标准模板 .xlsx。

- 输入:粘贴的 markdown 表、上传的 .xlsx、截图 OCR 后的文本 —— 任意形式
- 输出:`template1-cn.xlsx`(默认)/ `template1-en.xlsx`,18 列标准模板
- 图片:产品图**内嵌**进模板(不是链接文本),同一行多图会并排写入

> 与 [`template-resource-converter`](https://github.com/Edun1194/excel-template-resource-converter) 的区别:
> 那个 skill 部署了大量 vendor 适配器用 Python 脚本**批量**转固定品牌;
> 本 skill 让 **AI 自己读 + 转换**,适合单次 / 临时任务,或版式千奇百怪、来不及写适配器的场景。
> **不需要写任何解析脚本。**

两者产出的表**逐格一致**(模板逐字节相同、图片显示尺寸相同、C+ / 批发价语义相同),
所以同一批货单走哪条路都行,可以互换。

## 安装

```bash
git clone https://github.com/Edun1194/single-table-converter.git
cd single-table-converter
python3 -m pip install -r requirements.txt
```

也可以当 Mavis / Claude Code 的 skill 用 —— 目录结构就是 skill 标准布局
(`SKILL.md` + `references/` + `scripts/` + `assets/`)。

## 怎么用(给 AI 的用法)

把这几个要点告诉 agent 即可,细节在 `SKILL.md`:

1. **先整理成 `rows.json`**(list[dict],一行一个尺码),字段名见 `SKILL.md`
2. **再写进模板**:

```bash
python3 scripts/build_xlsx.py \
  --input rows.json \
  --template assets/templates/template1-cn.xlsx \
  --output out.xlsx
```

### rows.json 长这样

```json
[
  {
    "brand": "CANADA GOOSE",
    "sku": "1744M-9061",
    "description": "男士黑色红标连帽羽绒服",
    "category": "Down Jacket",
    "material": "",
    "color": "黑色",
    "color_code": "9061",
    "gender": "Man",
    "season": "FW25",
    "origin": "Canada",
    "cost": 860.0,
    "cplus_text": 18,
    "supply_price": null,
    "retail_price": 1500.0,
    "size": "S",
    "quantity": 5,
    "vendor_id": "HK",
    "image_path": "/abs/path/img1.jpg"
  }
]
```

| 字段 | 说明 |
|---|---|
| `cplus_text` | 整数 `18` / `-18`;**没给就写 `""`**(留空,别填 0)。C+ 是每单单独给的,不要当品牌默认值 |
| `supply_price` | 留 `null` 即可,脚本会用 `cost × (1 + C+/100)` 自动算 |
| `vendor_id` | 云仓列;旧键 `warehouse_id` 也兼容 |
| `image_path` / `image_paths` | 单图 / 多图路径(相对路径按 `rows.json` 所在目录解析) |

### CLI 参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `--input` | 必填 | rows.json 路径 |
| `--output` | 必填 | 输出 xlsx 路径 |
| `--template` | 内置 | 目标模板路径 |
| `--language {cn,en}` | `cn` | 模板语言 |
| `--image-row-height` | `92` | 有图数据行的行高(pt),图片等比缩到不超这个高度 |
| `--image-col-width` | `40` | 图片列列宽 |
| `--keep-zero-qty` | 关 | 保留数量 = 0 的行(默认跳过) |

## 模板字段(18 列)

```
品牌 | 图片 | 货号 | 描述 | 类目 | 材质 | 性别 | 颜色 | 色号
季节 | 产地 | 零售价 | cost | C+- | 批发价 | 尺码 | 数量 | 云仓
```

- `C+-` 存**纯整数**(`18` / `-18`),没有就留空
- `批发价` = `cost × (1 + C+-/100)`,没有 C+ 就留空(**不填 0**)
- `云仓` 是供应商 / 仓库 ID
- 图片在 **B 列**(col 2),显示高度 = 行高 92pt,**原图字节不动**(清晰度保留)

## 几条容易踩的规则

- **多尺码必须拆行** —— 否则下游按 (货号+颜色+尺码) 去重只留首行,数据丢光
- **C+ 没有就留空** —— 不要填 0
- **不要造 `TOTAL` 尺码** —— 源文件只有总数时,尺码留空、数量填总数
- **qty = 0 的行不输出** —— 期货单用 `null` 表示「还没填」,该行要保留
- **颜色空就空** —— 源文件没字段就别自己推断

更多见 `SKILL.md` 与 `references/app-parity.md`(与 app 端的逐条一致性对照)。

## 目录结构

```
.
├── SKILL.md                      # skill 主文档(算法、字段表、10 条通用坑)
├── README.md
├── requirements.txt
├── assets/templates/             # template1-cn.xlsx / template1-en.xlsx
├── scripts/
│   └── build_xlsx.py             # rows.json → 标准模板
├── examples/
│   └── sample-rows.json
└── references/
    ├── app-parity.md             # 与 app 端的一致性对照表
    ├── typical-formats.md        # 典型输入版式示例
    └── keywords.md               # 类目 / 性别 / 颜色 推断关键字
```

## License

[Apache License 2.0](LICENSE) —— Copyright 2026 Edun

允许商用、修改、再分发;使用者需保留版权与许可声明,需注明修改,并享有显式专利授权。
