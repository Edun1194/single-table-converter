---
name: single-table-converter
description: |
  单次表格转换:用户给一份 SKU 报价单/库存表(粘贴 markdown / 上传 xlsx / 截图),
  AI 自己识别版式,一次性输出标准模板(template1-cn 或 template1-en)格式的 .xlsx。
  不依赖多 vendor 适配器框架,适合在任意 LLM agent 中部署后即用。
  触发:用户给一份报价单 / 库存表 / 拉单,要求按模板整理 / 转模板 / 转成统一格式。
displayNames:
  zh-Hans: "单次表格转换"
---

# Single Table Converter

让 agent 在 **一次对话内**,把用户给的一份 SKU 表(任何形式)转成标准模板。

> 与 `template-resource-converter` 不同:那个是部署大量 vendor 适配器用 Python 脚本批量转;
> 本 skill 让 **LLM 自己读取 + 转换**,适合单次 / 临时任务、表格小、不值得写脚本的场景。

## 输入 (LLM 接收的形式)

agent 能识别任意一种:

1. **粘贴文本 / markdown 表格**
   ```
   SKU    名称           颜色     XS  S  M  L  XL
   1234M  男士黑色羽绒服  黑色     -   5  5  5  -
   ```
2. **上传的 .xlsx / .xls / .csv**
3. **截图 / 图片 OCR 后的文本**

无论哪种,**先在工作区里把表整理成 row-by-row 的 Python list[dict]**(见下)。

## 输出 (标准模板)

按 `assets/templates/template1-cn.xlsx`(默认)或 `template1-en.xlsx`,把整理后的 rows 写成 .xlsx。

模板 18 列,字段顺序固定(1-idx,**2026-10-19 起新模板**):

| #  | 中文 | 英文 | 类型 | 说明 |
|----|------|------|------|------|
| 1  | 品牌 | Brand           | str   | 品牌名,如 CANADA GOOSE |
| 2  | 图片 | Image           | bytes | 嵌入图片(优先);无图留空 |
| 3  | 货号 | International code | str | SKU / Style code;若源文件 STYLE 不含颜色代码,自动拼 `{STYLE}-{颜色代码}` 保证唯一 |
| 4  | 描述 | Described        | str   | 商品中文 / 英文描述 |
| 5  | 类目 | Category        | str   | Down Jacket / Vest / Sweater 等(可空) |
| 6  | 材质 | Material        | str   | 100% Cotton / Leather 等(可空) |
| 7  | 性别 | Gender          | str   | Man / Woman / Unisex / Kid(可空) |
| 8  | 颜色 | Color           | str   | 颜色名(可空,空时靠货号后缀保证唯一) |
| 9  | 色号 | Color Code      | str   | 颜色 code(可空) |
| 10 | 季节 | Season          | str   | 25FW / 26SS 等(可空) |
| 11 | 产地 | Made in         | str   | Italy / France / Canada 等(可空) |
| 12 | 零售价 | Retail price  | float | 建议零售价(可空) |
| 13 | cost | cost            | float | 出厂价/采购成本(必需) |
| 14 | C+-  | C+-             | int   | 折扣率**纯整数**如 `18` / `-18`(可空) |
| 15 | 批发价 | Wholesale price | float | `cost × (1 + C+-/100)`;**无 C+ 时留空,不填 0** |
| 16 | 尺码 | Size            | str   | 单个尺码,每行一个(多尺码必须拆) |
| 17 | 数量 | Quantity        | int   | 该尺码库存数;`null` = 还没填(期货单),保留该行 |
| 18 | 云仓 | Vendor ID       | str   | 仓库/供应商代码(可空,如 GZH / HZ) |

> **与旧模板差异**:图片列 3→2、货号 2→3、性别 9→7 / 颜色 7→8、C+ 之后无独立"供货价"列直接是"批发价"、英文描述字段名 `Described`(拼写)、英文产地 `Made in`、英文色号 `Color Code`(大写 C)。

> **模板 = app 端现行模板**。`assets/templates/template1-{cn,en}.xlsx` 与
> `excel-app-code/assets/templates/` 下的文件**逐字节一致**(2026-10-29 核对)。
> 字段名/列序/图片列位置有任何变化,两边必须一起改。

## ⚠️ C+ 是「每单单独给」的(2026-10-28 用户确认)

> 原话:「是我单独给的,不用保存」

- 每次转换以**用户当次给的 C+ 值**为准(例 `C+ 10` = 10%,批发价 = cost × 1.10)。
- **不要**把某个值当品牌默认值写死,**不要**套 `0.18` 之类的默认加价率。
- 用户没给 C+ → C+ 列和批发价**都留空**,并跟用户说明。
- C+ 写**纯整数**:`18` / `-18`。不带 `+` 号、不带 `%`、不是字符串 `"18"`。
  (`build_xlsx.py` 会把 `C+18%` / `C-15%` / `0.18` 统一归一成整数,给整数最省事。)

## 算法 (按顺序执行)

### Step 1:识别版式

读输入,判断是哪种:

- **多 section 报价单**:每个 section 内含尺码矩阵 → 按 section 逐个处理
- **扁平表**:一行一尺码或一行一多尺码
- **图片描述**:每个 SKU 一图或一组图 → 用第几张(主图)

不确定时,优先按"产品→尺码→库存"维度展开成 N 行。

### Step 2:归一化字段

字段名按上面表头对齐。对每行做:

1. **品牌 Brand**:源文件有就取,没有就空 / 询问用户(不要凭"标题首字母"自己推)
2. **货号 International code**:
   - 优先用源文件的 STYLE / Style code
   - 若该 STYLE 不含颜色代码 → 自动拼 `{STYLE}-{颜色代码}`
   - 分隔符固定 `-`
3. **描述 Description**:直接用源文件原文(必要时 normalize 空白)
4. **类目 / 性别 / 颜色 / 产地**:从描述/关键字推断(参见 references/keywords.md)
5. **成本 Cost** / **零售价 Retail price**:float,如 `€1.234,56` → 先去千分位 `1234.56`
6. **C+-**:**纯整数**,正数 `18`、负数 `-18` 保留符号。源是 `18%` / `C+18` 等统一成 `18`;`C-15%` → `-15`
7. **批发价 Wholesale price** = `cost × (1 + C+-/100)`;**只在 C+ 有值时计算**,否则留空
   (`build_xlsx.py` 会自动算:只给 `cost` + `cplus_text` 就行,不用自己算;
   给了 `supply_price` 就以你给的为准)
8. **尺码 Size**:每个尺码一行;若有矩阵行(尺码在列方向)→ 展开成 N 行
9. **库存 Quantity**:
   - 现货单:该尺码实际库存数。**qty = 0 的行不输出**(与 app 端 `is_positive_qty` 一致)
   - 期货单:没有真实库存 → `quantity` 写 `null`(留空),该行**保留**,等客户填
   - ⚠️ **不要造 `TOTAL` 尺码行**:app 端踩过这个坑(把总数当尺码写进去,下游全错)。
     源文件只有总数、没有分尺码时,尺码留空,数量填总数

### Step 3:多尺码拆行

单 SKU 多尺码必须拆成多行,**否则下游按货号去重会丢数据**(踩过坑)。

源数据:
```
SKU 1234M  黑色  XS=2  S=5  M=10  L=8
```
展开:
```
SKU 1234M  黑色  XS  2
SKU 1234M  黑色  S   5
SKU 1234M  黑色  M   10
SKU 1234M  黑色  L   8
```

### Step 4:去重(必要时)

若源数据同一 (sku+颜色+尺码) 出现多次,合并 qty;保留首次出现行的图片。

### Step 5:写入 .xlsx

用 `scripts/build_xlsx.py`:

```bash
python3 scripts/build_xlsx.py \
  --input rows.json \
  --template assets/templates/template1-cn.xlsx \
  --output out.xlsx
```

`rows.json` 格式:
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
    "season": "",
    "origin": "Canada",
    "cost": 860.0,
    "cplus_text": 18,
    "supply_price": null,
    "retail_price": null,
    "size": "S",
    "quantity": 5,
    "vendor_id": "HK",
    "image_path": "/abs/path/img1.jpg",
    "image_paths": ["/abs/path/img1.jpg", "/abs/path/img2.jpg"]
  }
]
```

| 字段 | 说明 |
|---|---|
| `cplus_text` | 整数 `18` / `-18`;**没给就写 `""`**(留空,别填 0) |
| `supply_price` | 留 `null` 即可,`build_xlsx.py` 会用 `cost × (1 + C+/100)` 自动算 |
| `vendor_id` | 云仓列(app 端同名字段);旧键 `warehouse_id` 也兼容 |
| `image_path` | 单张主图路径(相对路径按 `rows.json` 所在目录解析) |
| `image_paths` | 多张图(可选),会并排写在同一行,间距 8px |

`scripts/build_xlsx.py` 会:
1. 载入目标模板 → 从 R2 开始写 rows
2. 有 `image_path` / `image_paths` 则嵌入图片(等比缩到行高 92pt,**原图字节不动**)
3. 自动算批发价(有 cost + C+ 时)
4. 跳过 `quantity = 0` 的行(期货单的 `null` 保留)
5. 输出 `out.xlsx`

### CLI 参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `--input` | 必填 | rows.json 路径 |
| `--output` | 必填 | 输出 xlsx 路径 |
| `--template` | 内置 | 目标模板路径;不传就按 `--language` 取内置的 |
| `--language {cn,en}` | `cn` | 模板语言 |
| `--image-row-height` | `92` | 有图数据行的行高(pt),图片等比缩到不超这个高度 |
| `--image-col-width` | `40` | 图片列列宽(Excel 字符单位) |
| `--keep-zero-qty` | 关 | 保留数量 = 0 的行(默认跳过,与 app 端 `is_positive_qty` 一致) |

## 字段推断规则 (references/keywords.md 摘要)

**类目**(`_infer_category`):

| 关键字 | 英文 |
|--------|------|
| 羽绒服 / down jacket | Down Jacket |
| POLO / polo | Polo |
| T恤 / t-shirt / tee | Tee |
| 卫衣 / sweatshirt / hoodie | Sweatshirt |
| 衬衫 / shirt | Shirt |
| 针织衫 / knit | Knitwear |
| 外套 / jacket | Jacket |
| 大衣 / coat | Coat |
| 马甲 / vest | Vest |
| 休闲裤 / pants / trousers | Pants |
| 风衣 / trench | Trench |
| 帽 / hat | Hat |

**性别**(`_infer_gender`):

| 关键字 | 值 |
|--------|-----|
| 男士 / men's / men | Man |
| 女士 / women's / women | Woman |
| 男 / male | Man |
| 女 / female | Woman |
| 中性 / unisex | Unisex |
| 童 / kid / kids / children | Kid |

**颜色**(`_infer_color`):黑色 / 白色 / 深蓝色 / 浅灰色 / 深灰色 / 灰色 / 蓝色 / 红色 / 绿色 / 米色 / 驼色 / 紫色 / 粉色 / 黄色 / 橘色 / 棕色 / 卡其 / 藏青 / 墨绿 / 酒红 / 象牙 / 杏色 / 军绿

**产地**:意大利品牌 → Italy / 法国品牌 → France / 加拿大(CANADA GOOSE 等)→ Canada;若能直接读源文件产地字段则优先。

## 通用坑(LLM 必看)

1. **货号不要用爱莎码 / 国内 SKU** — 用源文件的 STYLE / 国际码;不含颜色代码时拼 `{STYLE}-{颜色}`。
2. **多尺码必须拆行** — 否则下游系统按 (货号+颜色+尺码) 去重只留首行,数据丢光。
3. **C+ 没有就留空** — 不要填 `0` 或 `0.0`,会污染公式。
4. **supply_price 没有 C+ 就别算** — 留 `null`,不要写 `cost * 1.0`。
5. **颜色空就空** — 不要凭"标题首字母"或"前缀"自己推 M/W,源文件没字段就空。
6. **图片 anchor**:openpyxl 0-idx row = 1-idx row - 1;不同供应商偏移风格不同(怀远/MAX MARA 用 `r-1`,老 PRADA/欧洲用 `r-2`),尽量 try 多个 key 兼容。
7. **数字清洗**:`1.234,56` (欧式)→ 去点号替逗号 → `1234.56`。
8. **不要造 `TOTAL` 尺码** — 源文件只有总数时,尺码留空、数量填总数。app 端踩过这坑(把总数当尺码写进 Size 列)。
9. **qty = 0 的行不输出** — 现货单跟 app 端 `is_positive_qty` 一致;期货单用 `null` 保留行。
10. **图片显示高度 = 行高(92pt),等比** — 不要写死 `img.width = 280` 这种值:
    openpyxl 的 `Image.width/height` 是**像素**,写 `X * 9525` 会双倍乘(图变 79ft 高),
    写死 280×220 还会**拉伸变形**。`build_xlsx.py` 已按 app 端方式处理,别自己改。

## 与 app 端(app 侧 excel-app-code)的一致性

本 skill 的输出要和 app 端 `convert_vendor_quote.py` 的产物**可互换**:
模板逐字节一致、字段映射一致、图片显示尺寸一致。

逐条对照表见 `references/app-parity.md`(含 2026-10-29 实测的比对结果)。

## 部署到其他 agent

```bash
# 方式 1:git clone(推荐)
git clone https://github.com/Edun1194/single-table-converter.git \
  ~/.minimax/skills/single-table-converter/

# 方式 2:解压 zip 包
unzip single-table-converter-skill.zip -d ~/.minimax/skills/single-table-converter/
```

或作为独立 Python 包使用:
```bash
export PYTHONPATH="<skill-dir>/scripts:$PYTHONPATH"
python3 build_xlsx.py --input rows.json --template <skill-dir>/assets/templates/template1-cn.xlsx --output out.xlsx
```

## 依赖

```
openpyxl>=3.1
pillow>=10.0
```

## 一句话总结

**给 LLM 一份表,让它按这个 skill 转成标准模板 18 列 .xlsx。**
LLM 不写解析脚本 — 这次对话里直接读 → 整理 → 调 build_xlsx.py 输出。
