# 与 app 端的一致性对照(app-parity)

> 目标:本 skill 的输出与 app 端(`excel-app-code` 的 `convert_vendor_quote.py` + `convert_vendor_quote_shared/`)
> 产物**可互换** —— 同一批数据走哪条路,出来的表长得一样。
>
> 核对日期:2026-10-29 · 核对方式:同一批 rows 分别走两条路,逐格比对 + 图片 EMU 比对。

## 1. 模板

| 项 | 状态 |
|---|---|
| `assets/templates/template1-cn.xlsx` vs app `assets/templates/template1-cn.xlsx` | **逐字节一致**(17,026 bytes) |
| `assets/templates/template1-en.xlsx` vs app `assets/templates/template1-en.xlsx` | **逐字节一致**(17,886 bytes) |

18 列(中)/ 18 列(英,末尾多一个空列),列序:

```
品牌 | 图片 | 货号 | 描述 | 类目 | 材质 | 性别 | 颜色 | 色号
季节 | 产地 | 零售价 | cost | C+- | 批发价 | 尺码 | 数量 | 云仓
```

⚠️ **模板改了,两边必须一起改**(本 skill 的 + app 的)。
app 端模板在 `excel-app-code/assets/templates/`,是唯一事实来源。

## 2. 字段映射

`build_xlsx.py` 的 row 字段名与 app 端 `map_row_to_template()` 一致:

| 模板列 | 本 skill row key | app 端 row key | 一致 |
|---|---|---|---|
| 品牌 | `brand` | `brand` | ✅ |
| 货号 | `sku` | `sku` | ✅ |
| 描述 | `description` | `description` | ✅ |
| 类目 | `category` | `category` | ✅ |
| 材质 | `material` | `material` | ✅ |
| 性别 | `gender` | `gender`(可选) | ✅ |
| 颜色 | `color` | `color` | ✅ |
| 色号 | `color_code` | `color_code`(可选) | ✅ |
| 季节 | `season` | `season` | ✅ |
| 产地 | `origin` | `origin` | ✅ |
| 零售价 | `retail_price` | `retail_price` | ✅ |
| cost | `cost` | `cost` | ✅ |
| C+- | `cplus_text` | `cplus_text` | ✅ |
| 批发价 | `supply_price` | `supply_price` | ✅ |
| 尺码 | `size` | `size` | ✅ |
| 数量 | `quantity` | `quantity` | ✅ |
| 云仓 | `vendor_id`(兼容旧键 `warehouse_id`) | `vendor_id` | ✅ |

## 3. 图片

app 端实现:`convert_vendor_quote_shared/__init__.py` 的 `embed_images_in_row()`。

| 规则 | app 端 | 本 skill(2026-10-29 之前) | 本 skill(现在) |
|---|---|---|---|
| 锚点类型 | `OneCellAnchor` | `add_image(img, "B{row}")`(隐式 anchor) | `OneCellAnchor` ✅ |
| 显示尺寸 | 原图像素 × 9525 EMU | **写死 280×220 px** ❌ | 原图像素 × 9525 EMU ✅ |
| 等比缩放 | `cy > 行高 EMU` 时按比例缩 | ❌ 无(写死值,且拉伸变形) | ✅ |
| 高度上限 | 行高 92pt = 1,168,400 EMU | ❌(210pt,超出行高 2.3 倍) | ✅ 92pt |
| 原图字节 | 不动(清晰度保留) | 不动 | 不动 ✅ |
| 同一行多图 | 按显示宽度累加 `colOff`,间距 8px | 不支持 | ✅ 支持 `image_paths` |

**实测比对**(800×1000 的图 + 92pt 行高):

| | anchor row(0-idx) | 显示尺寸 |
|---|---|---|
| app 端 | 1 / 2 / 3 | 73.60pt × 92.00pt |
| 本 skill(旧) | 1 / 2 / 3 | 210.00pt × 165.00pt ❌ |
| 本 skill(新) | 1 / 2 / 3 | **73.60pt × 92.00pt** ✅ |

> ⚠️ **openpyxl 的 `Image.width` / `Image.height` 是像素,不是 EMU**。
> 写 `img.width = 92` → 实际 ext `cx = 92 × 9525` EMU(= 92pt)。
> 写 `img.width = 92 * 9525` → **双倍乘**,图会变成 79ft 高。

## 4. C+ 与批发价

| 规则 | app 端 | 本 skill |
|---|---|---|
| C+ 类型 | vendor 自带 → **int**;`--cplus` 路径 → `str` | 统一 **int** ✅ |
| C+ 归一 | 调用方给什么写什么 | `C+18%` / `C-15%` / `0.18` → 整数 |
| 没有 C+ | 留空 | 留空 ✅ |
| 批发价 | dispatcher 层算 `cost × (1 + C+/100)`,round 4 | `build_xlsx.py` 自动算,round 2 ✅ |
| 没有 C+ 时批发价 | 留空(不填 0) | 留空 ✅ |

> 📌 **app 端的一个小瑕疵**(本 skill 不复刻):
> `convert_vendor_quote.py` 的 `--cplus` 路径写的是 `row["cplus_text"] = str(cplus_int)`,
> 落盘是**文本** `"10"`;而 vendor 自带的 C+ 落盘是整数。
> 实测历史产物 905 个 C+ 单元格**全是 int**,只有 `--cplus` 这条新路径是文本。
> 本 skill 统一写 int(符合两个 SKILL.md 都写明的「纯整数」规则)。
> 若你决定改 app 端 → `convert_vendor_quote.py` 里把 `str(cplus_int)` 改成 `int(cplus_int)` 即可。

实测批发价精度:app 端历史产物最多 2 位小数(`473.55`),round 2 / round 4 无差别。

## 5. 行数过滤

| 场景 | app 端 | 本 skill |
|---|---|---|
| 现货单 `quantity = 0` | 跳过(`is_positive_qty`) | 跳过 ✅(`--keep-zero-qty` 可关) |
| 期货单 `quantity = null` | 保留(适配器自己决定) | 保留 ✅ |

## 6. 行高

| 规则 | app 端 | 本 skill |
|---|---|---|
| 有图行 | `image_row_height`(默认 92pt) | 92pt ✅ |
| 无图行 | 沿用源行高 / 模板默认 | 沿用模板默认 ✅ |

## 7. 还没对齐的(知道就行)

- app 端支持**图片目录批量匹配**(`image-directory-matching`)、**PDF 解析**、**OLE2 .xls 抽图** —— 本 skill 不做(LLM 一次性读图)。
- app 端有 **C+ 规则文件 / 行级覆盖**(`CplusResolver` 3 层优先级)—— 本 skill 只有「用户当次给一个值」。
- app 端有 **media 去重**(同图多尺码行只写一份 media)—— 本 skill 单次任务行数少,没做。
