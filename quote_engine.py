# -*- coding: utf-8 -*-
"""
天航供应链 自动报价引擎
从《天航空派同行代理价.xlsx》读取价格，解析询价文本，计算计费重并输出各渠道报价。

规则依据（与销售确认）：
- 材积重 = 长*宽*高 / 6000
- 单件不足 10KG 按 10KG 计费；单票最低 12KG
- 21KG+ 分泡：计费重 = (泡重 - 实重) * (1 - 分泡比) + 实重（分泡60% 即 *0.4）
- 100KG+ 密比折扣：1:300 减2 / 1:500 减3 / 1:800 减4（元/KG）
- 暂停/单询渠道不展示
"""
import re
import sys
import os
import json
import math
import openpyxl

# 价格表默认与代码放在同一目录；也可用环境变量 PRICE_XLSX 指定绝对路径
XLSX_PATH = os.environ.get("PRICE_XLSX") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "天航空派同行代理价11.7.xlsx")

# ---------------- 报价表加载 ----------------

# 美国空派渠道在 sheet 中的定位行（渠道名所在行，价格在其下 1~3 行）
US_AIR_BLOCKS = [
    {"channel": "美国空派B",   "header_row": 4,  "data_rows": [5, 6, 7],  "zones": {"美西": "8-9", "美中": "4-7", "美东": "0-3"}},
    {"channel": "美国空派A",   "header_row": 12, "data_rows": [13, 14, 15], "zones": {"美西": "8-9", "美中": "4-7", "美东": "0-3"}},
    {"channel": "美国空派ORD", "header_row": 20, "data_rows": [21], "zones": {"全美": "0-9"}},
    {"channel": "美国空派JFK", "header_row": 27, "data_rows": [28], "zones": {"全美": "0-9"}},
]
TIERS = ["12KG+", "21KG+", "51KG+", "101KG+"]  # 对应 D/E/F/G 列


def load_us_air_prices(xlsx_path=XLSX_PATH):
    """返回 {channel: {zone: {tier: price_or_str}}, ...}，暂停/单询返回字符串"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["美国空派 "]
    table = {}
    for blk in US_AIR_BLOCKS:
        ch = {}
        for zone_name, r in zip(blk["zones"].keys(), blk["data_rows"]):
            prices = {}
            for i, col in enumerate("DEFG"):
                v = ws[f"{col}{r}"].value
                prices[TIERS[i]] = v
            ch[zone_name] = prices
        table[blk["channel"]] = {
            "zones": ch,
            "时效": ws[f"H{blk['header_row'] + 1}"].value or "",
            "备注": (ws[f"I{blk['header_row'] + 1}"].value or "").strip(),
        }
    return table


def load_us_small_prices(xlsx_path=XLSX_PATH):
    """美国空派小货：重量(KG)->运费(元/票)，0.5KG 步进"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["美国小货"]
    prices = {}
    for row in ws.iter_rows(min_row=5, min_col=2, max_col=3):
        w, p = row[0].value, row[1].value
        if isinstance(w, (int, float)) and isinstance(p, (int, float)):
            prices[float(w)] = float(p)
    return prices


# ---------------- 附加费规则 ----------------

OVERSIZE_RULES = [  # 返回 (说明, 金额/件) 或 None；按件叠加
    lambda L, W, H, girth, w: ("单边120-200CM", 180) if 120 <= max(L, W, H) <= 200 else None,
    lambda L, W, H, girth, w: ("单边超120且第二边超75", 180) if (max(L, W, H) > 120 and sorted([L, W, H])[-2] > 75) else None,
    lambda L, W, H, girth, w: ("单边未超120但另两边超75", 180) if (max(L, W, H) <= 120 and sorted([L, W, H])[-2] > 75) else None,
    lambda L, W, H, girth, w: ("围长260-320CM", 180) if 260 <= girth < 320 else None,
    lambda L, W, H, girth, w: ("围长320-410CM", 1040) if 320 <= girth <= 410 else None,
    lambda L, W, H, girth, w: ("单件实重超22KG", 180) if 22 < w <= 40 else None,
]

PRODUCT_SURCHARGE_KEYWORDS = [  # (关键词, +元/KG)。正式版接入"美国产品附加费"HTS表
    (["木质", "竹", "牙刷", "小剪刀", "圆珠笔", "铅笔", "服装"], 1),
    (["灯具", "电器", "风扇", "手电筒", "电线", "铜", "汽车配件", "摩托", "不锈钢", "金属"], 2),
    (["草编", "钢制螺栓", "螺母"], 3),
    (["假发", "扑克", "骰子", "面罩", "创可贴", "防护服", "体温计", "玩具", "婴儿", "宠物玩具", "儿童"], 5),
]


def product_surcharge(product_desc: str):
    desc = product_desc or ""
    hits = []
    for kws, rate in PRODUCT_SURCHARGE_KEYWORDS:
        if any(k in desc for k in kws):
            hits.append(rate)
    return max(hits) if hits else 0  # 同产品多档取最高


# ---------------- 计费重计算 ----------------

def chargeable_weight(pieces, dims, weight_per_piece, min_piece=10.0, min_ticket=12.0,
                      split_bulge_ratio=0.6):
    """
    pieces: 件数; dims: (长,宽,高) cm 每件相同; weight_per_piece: 单件实重 kg
    返回 (计费重, 明细dict)
    """
    L, W, H = dims
    vol_per_piece = L * W * H / 6000.0
    # 单件不足10KG按10KG计费后的"实重"基数
    actual_base_per_piece = max(weight_per_piece, min_piece)
    actual_total = actual_base_per_piece * pieces
    vol_total = vol_per_piece * pieces  # 泡重
    detail = {
        "单件材积重": round(vol_per_piece, 2),
        "单件计费基数": round(actual_base_per_piece, 2),
        "泡重合计": round(vol_total, 2),
        "实重合计": round(actual_total, 2),
    }
    if vol_total > actual_total:
        # 泡货：21KG+ 分泡（分泡60% -> 乘 0.4）
        charge = (vol_total - actual_total) * (1 - split_bulge_ratio) + actual_total
        detail["分泡方式"] = f"分泡{int(split_bulge_ratio*100)}%"
    else:
        charge = actual_total
        detail["分泡方式"] = "实重计费（无泡可分）"
    charge = max(charge, min_ticket)  # 单票最低12KG
    detail["计费重"] = round(charge, 2)
    return charge, detail


def density_discount(charge_weight):
    """100KG+ 密比折扣：1:300减2 / 1:500减3 / 1:800减4"""
    density = charge_weight * 6000  # 相对1CBM
    # 密比= 计费重/体积；此处按行业惯例用 kg 与体积比判断，简化：以每kg对应cm3
    # 实际口径：材积比 = 6000*体积 / 计费重
    return None  # 在 quote() 中结合体积计算


# ---------------- 询价文本解析 ----------------

NUM = r"(\d+(?:\.\d+)?)"

def parse_inquiry(text: str):
    """
    从询价文本提取：件数、尺寸、单件重量、邮编、品名、渠道偏好
    兼容 "39*28*25cm 17kg/箱 2箱" 等写法
    """
    info = {"pieces": 1, "dims": None, "weight": None, "zip": None,
            "product": None, "channel_hint": None, "address_type": None}

    m = re.search(NUM + r"\s*[*×xX]\s*" + NUM + r"\s*[*×xX]\s*" + NUM, text)
    if m:
        info["dims"] = tuple(float(g) for g in m.groups())

    # 重量：17kg/箱 或 17kg
    m = re.search(NUM + r"\s*(?:kg|KG|公斤)", text)
    if m:
        info["weight"] = float(m.group(1))

    # 件数：2箱 / 5件 / 1票5件；排除"7件套""3件装"这类商品件数表述
    m = re.search(r"(\d+)\s*箱", text)
    if not m:
        m = re.search(r"(\d+)\s*件(?!\s*[套装])", text)
    if m:
        info["pieces"] = int(m.group(1))
    m = re.search(r"(\d+)\s*票\s*(\d+)\s*件", text)
    if m:
        info["pieces"] = int(m.group(2))

    # 邮编（美国5位）
    m = re.search(r"\b(\d{5})\b", text)
    if m:
        info["zip"] = m.group(1)

    # 地址类型
    if re.search(r"FBA|亚马逊", text, re.I):
        info["address_type"] = "FBA"
    elif re.search(r"住宅|residential", text, re.I):
        info["address_type"] = "住宅"
    elif re.search(r"商业|公司|commercial", text, re.I):
        info["address_type"] = "商业"

    # 渠道偏好
    for ch in ["空派B", "空派A", "空派ORD", "空派JFK", "小货", "空卡"]:
        if ch in text:
            info["channel_hint"] = "美国空派" + ch[-1] if ch.startswith("空派") else ch
            break

    # 品名：去掉数字和单位后剩余中文片段（简化，正式版用大模型识别图片+文本）
    m = re.search(r"([\u4e00-\u9fa5]{2,10}?)(?:\s|，|,|$|\d)", text)
    cleaned = re.sub(r"[\d.*×xX/\\ ]", " ", text)
    words = re.findall(r"[\u4e00-\u9fa5]{2,6}", cleaned)
    info["product_guess"] = " ".join(words[:3]) if words else None
    return info


# ---------------- 报价 ----------------

def quote_us_air(text: str, xlsx_path=XLSX_PATH):
    table = load_us_air_prices(xlsx_path)
    info = parse_inquiry(text)
    if not info["dims"] or not info["weight"]:
        return {"ok": False, "msg": "缺少尺寸或重量，请补充（例：39*28*25cm 17kg/箱 2箱 邮编+地址）"}

    pieces = info["pieces"]
    L, W, H = info["dims"]
    girth = (L + W) * 2  # 围长 = 最长边+底周/2，此处按 (长+宽)*2 简化，单件重按最大边计
    cw, wdetail = chargeable_weight(pieces, info["dims"], info["weight"])

    # 密比折扣（100KG+）
    density = cw / (L * W * H * pieces / 1_000_000) if L * W * H else 999
    discount = 0
    if cw >= 100:
        ratio = cw / (L * W * H * pieces / 6000 / 100)  # 相对材积比
        # 口径：材积/计费重（即 1:xxx）
        vol = L * W * H * pieces / 6000.0
        dens_ratio = cw / vol if vol else 1
        if dens_ratio >= 800: discount = 4
        elif dens_ratio >= 500: discount = 3
        elif dens_ratio >= 300: discount = 2

    # 单件超重/超长附加（按件）
    oversize = []
    for rule in OVERSIZE_RULES:
        r = rule(L, W, H, girth, info["weight"])
        if r:
            oversize.append(r)
    oversize_fee = sum(f for _, f in oversize) * pieces

    # 产品附加
    ps_rate = product_surcharge(info.get("product_guess") or "")
    ps_fee = ps_rate * cw

    # 分区
    zone = "全美"
    if info["zip"]:
        first = int(info["zip"][0])
        if first >= 8: zone = "美西"
        elif first >= 4: zone = "美中"
        else: zone = "美东"

    results = []
    for ch, data in table.items():
        z = data["zones"].get(zone) or data["zones"].get("全美")
        if not z:
            continue
        tier = "12KG+"
        if cw >= 101: tier = "101KG+"
        elif cw >= 51: tier = "51KG+"
        elif cw >= 21: tier = "21KG+"
        unit = z.get(tier)
        if not isinstance(unit, (int, float)):  # 暂停/单询 -> 不展示
            continue
        unit_eff = unit - discount
        freight = unit_eff * cw
        total = freight + oversize_fee + ps_fee
        results.append({
            "渠道": ch, "时效": data["时效"], "单价": round(unit_eff, 1),
            "计费重": cw, "运费": round(freight, 0),
            "附加费": round(oversize_fee + ps_fee, 0), "合计": round(total, 0),
        })
    results.sort(key=lambda x: x["合计"])

    return {
        "ok": True,
        "输入": info,
        "计费明细": wdetail,
        "计费重": cw,
        "分区": zone,
        "密比折扣": f"-{discount}元/KG" if discount else "无",
        "超长超重附加": oversize if oversize else "无",
        "产品附加费": f"+{ps_rate}元/KG" if ps_rate else "无（请核实材质/品类）",
        "报价": results,
        "提醒": [
            "偏远费2元/KG（最低30元/件）待邮编底表/UPS官方表确认后计算",
            "非FBA地址+1元/KG；地址类型未识别时未计入，请核实",
        ],
    }


def format_reply(result):
    if not result["ok"]:
        return result["msg"]
    lines = ["📦 自动报价（代理价，未含税点）", ""]
    d = result["计费明细"]
    lines.append(f"计费重：{result['计费重']}KG（{d['单件材积重']}材积/件，{d['分泡方式']}），分区：{result['分区']}")
    if result["超长超重附加"] != "无":
        lines.append("超长/超重：" + "、".join(f"{n}+{f}/件" for n, f in result["超长超重附加"]))
    if result["产品附加费"] != "无（请核实材质/品类）":
        lines.append(f"产品附加：{result['产品附加费']}")
    lines.append("")
    lines.append("渠道 | 时效 | 单价 | 运费+附加 | 合计")
    lines.append("--- | --- | --- | --- | ---")
    for r in result["报价"]:
        lines.append(f"{r['渠道']} | {r['时效']} | {r['单价']}元/KG | {r['运费']:.0f}+{r['附加费']:.0f} | **{r['合计']:.0f}元**")
    lines.append("")
    for tip in result["提醒"]:
        lines.append("⚠️ " + tip)
    return "\n".join(lines)


if __name__ == "__main__":
    text = sys.argv[1] if len(sys.argv) > 1 else "量勺7件套 39*28*25cm 17kg/箱 2箱 美国 Halton City, TX 76117 空运附加有哪些"
    result = quote_us_air(text)
    print(format_reply(result))
