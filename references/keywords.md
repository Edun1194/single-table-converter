# 字段推断规则 (Field Inference)

LLM 在没有显式字段时,从描述/标题/关键字推断类目、性别、颜色、产地。

## 类目 (Category)

**判断顺序**:精确 → 大类 → 留空。不要凭"首字母"或"风格"自创。

| 中文 | 英文 | 备注 |
|------|------|------|
| 羽绒服 / down jacket | Down Jacket | |
| POLO / polo | Polo | 大写匹配 |
| T恤 / t恤 / t-shirt / tee | Tee | |
| 卫衣 / 套头 / hoodie / sweatshirt | Sweatshirt | |
| 衬衫 / shirt | Shirt | |
| 针织衫 / 毛衣 / knit / sweater | Knitwear | |
| 外套 / jacket | Jacket | |
| 大衣 / coat | Coat | |
| 马甲 / 背心 / vest | Vest | |
| 休闲裤 / 长裤 / pants / trousers | Pants | |
| 风衣 / trench | Trench | |
| 帽 / hat | Hat | |
| 鞋 / shoes / sneakers / boots | Shoes | |
| 包 / bag | Bag | |
| 围巾 / scarf | Scarf | |
| 手套 / gloves | Gloves | |
| 腰带 / 皮带 / belt | Belt | |

## 性别 (Gender)

**优先级**:长关键字 > 短关键字(如 `男士` 优先于 `男`)。

| 中文 | 英文 |
|------|------|
| 男士 / men's / men | Man |
| 男 / male | Man |
| 女士 / women's / women | Woman |
| 女 / female | Woman |
| 中性 / unisex | Unisex |
| 童 / kid / kids / children / 儿童 | Kid |
| 婴儿 / baby / infant | Kid |

## 颜色 (Color)

22 色标准列表,源文件出现的就取,没就空。

```python
COLOR_KEYWORDS = [
    "黑色", "白色", "深蓝色", "浅灰色", "深灰色", "灰色",
    "蓝色", "红色", "绿色", "米色", "驼色", "紫色", "粉色",
    "黄色", "橘色", "棕色", "卡其", "藏青", "墨绿", "酒红",
    "象牙", "杏色", "军绿",
]
```

英文:Black / White / Dark Blue / Light Grey / Dark Grey / Grey / Blue / Red / Green / Beige / Camel / Purple / Pink / Yellow / Orange / Brown / Khaki / Navy / Dark Green / Burgundy / Ivory / Apricot / Army Green

## 产地 (Origin)

源文件有 `产地` 字段优先;没有就按品牌典型产地:

| 品牌 | 产地 |
|------|------|
| CANADA GOOSE | Canada |
| MOORER | Italy |
| PRADA | Italy |
| GUCCI | Italy |
| MONCLER | France |
| BURBERRY | UK |
| LONGCHAMP | France |
| POLO RALPH LAUREN | USA |
| TORY BURCH | USA |
| COACH | USA |
| HOKA | France |

无品牌/无明确线索 → 留空。

## 价格清洗

**欧式**:`1.234,56` → 去千分位 + 替换逗号 → `1234.56`

```python
def parse_price(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("€", "").replace("$", "").strip()
    # 欧式:点号是千分位,逗号是小数
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None
```

**货币符号**:€ / USD$ / £ / ¥ — 全部剥掉;价格本身判断单位。

## C+ / 折扣 解析

源 → 统一存为**纯数字字符串**(`18` / `-18`):

| 源文本 | 存入 |
|--------|------|
| `C+18%` | `18` |
| `c+18` | `18` |
| `C-15%` | `-15` |
| `C+ 18` | `18` |
| `18` | `18` |
| `-18` | `-18` |
| (空) | `` (留空,千万别填 0) |

**supply_price** 计算:
```python
if cplus_text:  # 有值才算
    p = float(cplus_text)
    supply = round(cost * (1 + p / 100), 2)
else:
    supply = None  # 空
```
