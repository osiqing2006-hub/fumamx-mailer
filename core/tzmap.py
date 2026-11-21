"""国家 → 时区换算（2026-10-05 加）。

背景
----
实测反馈「发送的当地时间和我设置的不一样」。查日志确认：

    23:05:49 [定时] 等待您在弹窗里点「是」…（共 119 位，当地时间 2026/10/06 09:00）
    23:06:10   定时-当地时间: 2026/10/06 09:00     ← 英国客户
    ...
    23:22:26   定时-当地时间: 2026/10/06 12:53     ← 美国客户

用户在④页填的 09:00 被**当成各客户自己的当地时间**写进孚盟：
英国当地 09:00 = 北京 16:00，美国当地 09:00 = 北京 22:00。
所以用户以为"10/6 早上 9 点发"，实际各国发作息差 6~12 小时。

本模块让用户可以选「按北京时间定」：填北京时间，脚本按客户国家时区
换算成当地时间再写进孚盟；同时把两个时间都打进日志，一眼可核对。

实现要点
--------
· 优先用标准库 zoneinfo（自动处理夏令时，venv 里已装 tzdata，598 个时区）；
· 拿不到 tzdata 时回落内置的固定偏移表（按北半球夏令时经验值，精度够用）；
· 认不出的国家返回 None，调用方保持原值不换算（宁可不换，也不换错）。
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

CN = timezone(timedelta(hours=8))

# 中文国家名 → IANA 时区（代表城市/首都）。多时区国家取最主要商业区。
# 覆盖外贸常见国家；不在表里的返回 None，不做猜测。
_TZ: dict[str, str] = {
    # ---- 亚洲 ----
    "中国": "Asia/Shanghai", "中国大陆": "Asia/Shanghai",
    "中国香港": "Asia/Hong_Kong", "香港": "Asia/Hong_Kong",
    "中国澳门": "Asia/Macau", "澳门": "Asia/Macau",
    "中国台湾": "Asia/Taipei", "台湾": "Asia/Taipei",
    "日本": "Asia/Tokyo",
    "韩国": "Asia/Seoul", "南韩": "Asia/Seoul", "韩国（南）": "Asia/Seoul",
    "朝鲜": "Asia/Pyongyang",
    "蒙古": "Asia/Ulaanbaatar",
    "新加坡": "Asia/Singapore",
    "马来西亚": "Asia/Kuala_Lumpur",
    "泰国": "Asia/Bangkok",
    "越南": "Asia/Ho_Chi_Minh",
    "柬埔寨": "Asia/Phnom_Penh",
    "老挝": "Asia/Vientiane",
    "缅甸": "Asia/Yangon",
    "菲律宾": "Asia/Manila",
    "印度尼西亚": "Asia/Jakarta", "印尼": "Asia/Jakarta",
    "文莱": "Asia/Brunei",
    "东帝汶": "Asia/Dili",
    "印度": "Asia/Kolkata",
    "巴基斯坦": "Asia/Karachi",
    "孟加拉国": "Asia/Dhaka", "孟加拉": "Asia/Dhaka",
    "斯里兰卡": "Asia/Colombo",
    "尼泊尔": "Asia/Kathmandu",
    "不丹": "Asia/Thimphu",
    "马尔代夫": "Asia/Male",
    "阿富汗": "Asia/Kabul",
    "伊朗": "Asia/Tehran",
    "伊拉克": "Asia/Baghdad",
    "叙利亚": "Asia/Damascus",
    "黎巴嫩": "Asia/Beirut",
    "约旦": "Asia/Amman",
    "以色列": "Asia/Jerusalem",
    "巴勒斯坦": "Asia/Hebron",
    "沙特阿拉伯": "Asia/Riyadh", "沙特": "Asia/Riyadh",
    "科威特": "Asia/Kuwait",
    "卡塔尔": "Asia/Qatar",
    "巴林": "Asia/Bahrain",
    "阿联酋": "Asia/Dubai", "阿拉伯联合酋长国": "Asia/Dubai", "迪拜": "Asia/Dubai",
    "阿曼": "Asia/Muscat",
    "也门": "Asia/Aden",
    "土耳其": "Europe/Istanbul",
    "塞浦路斯": "Asia/Nicosia",
    "哈萨克斯坦": "Asia/Almaty",
    "乌兹别克斯坦": "Asia/Tashkent",
    "吉尔吉斯斯坦": "Asia/Bishkek",
    "塔吉克斯坦": "Asia/Dushanbe",
    "土库曼斯坦": "Asia/Ashgabat",
    "格鲁吉亚": "Asia/Tbilisi",
    "亚美尼亚": "Asia/Yerevan",
    "阿塞拜疆": "Asia/Baku",
    # ---- 欧洲 ----
    "英国": "Europe/London", "英格兰": "Europe/London", "苏格兰": "Europe/London",
    "爱尔兰": "Europe/Dublin",
    "法国": "Europe/Paris",
    "德国": "Europe/Berlin",
    "荷兰": "Europe/Amsterdam",
    "比利时": "Europe/Brussels",
    "卢森堡": "Europe/Luxembourg",
    "瑞士": "Europe/Zurich",
    "奥地利": "Europe/Vienna",
    "意大利": "Europe/Rome",
    "西班牙": "Europe/Madrid",
    "葡萄牙": "Europe/Lisbon",
    "希腊": "Europe/Athens",
    "丹麦": "Europe/Copenhagen",
    "瑞典": "Europe/Stockholm",
    "挪威": "Europe/Oslo",
    "芬兰": "Europe/Helsinki",
    "冰岛": "Atlantic/Reykjavik",
    "波兰": "Europe/Warsaw",
    "捷克": "Europe/Prague",
    "斯洛伐克": "Europe/Bratislava",
    "匈牙利": "Europe/Budapest",
    "斯洛文尼亚": "Europe/Ljubljana",
    "克罗地亚": "Europe/Zagreb",
    "波黑": "Europe/Sarajevo",
    "塞尔维亚": "Europe/Belgrade",
    "黑山": "Europe/Podgorica",
    "北马其顿": "Europe/Skopje", "马其顿": "Europe/Skopje",
    "阿尔巴尼亚": "Europe/Tirane",
    "保加利亚": "Europe/Sofia",
    "罗马尼亚": "Europe/Bucharest",
    "摩尔多瓦": "Europe/Chisinau",
    "爱沙尼亚": "Europe/Tallinn",
    "拉脱维亚": "Europe/Riga",
    "立陶宛": "Europe/Vilnius",
    "白俄罗斯": "Europe/Minsk",
    "乌克兰": "Europe/Kyiv",
    "俄罗斯": "Europe/Moscow",
    # ---- 美洲 ----
    "美国": "America/New_York", "美利坚合众国": "America/New_York",
    "加拿大": "America/Toronto",
    "墨西哥": "America/Mexico_City",
    "巴西": "America/Sao_Paulo",
    "阿根廷": "America/Argentina/Buenos_Aires",
    "智利": "America/Santiago",
    "秘鲁": "America/Lima",
    "哥伦比亚": "America/Bogota",
    "委内瑞拉": "America/Caracas",
    "厄瓜多尔": "America/Guayaquil",
    "玻利维亚": "America/La_Paz",
    "巴拉圭": "America/Asuncion",
    "乌拉圭": "America/Montevideo",
    "哥斯达黎加": "America/Costa_Rica",
    "巴拿马": "America/Panama",
    "危地马拉": "America/Guatemala",
    "洪都拉斯": "America/Tegucigalpa",
    "萨尔瓦多": "America/El_Salvador",
    "尼加拉瓜": "America/Managua",
    "多米尼加": "America/Santo_Domingo",
    "牙买加": "America/Jamaica",
    "古巴": "America/Havana",
    "波多黎各": "America/Puerto_Rico",
    "特立尼达和多巴哥": "America/Port_of_Spain",
    "海地": "America/Port-au-Prince",
    # ---- 大洋洲 ----
    "澳大利亚": "Australia/Sydney", "澳洲": "Australia/Sydney",
    "新西兰": "Pacific/Auckland",
    "巴布亚新几内亚": "Pacific/Port_Moresby",
    "斐济": "Pacific/Fiji",
    "萨摩亚": "Pacific/Apia",
    "汤加": "Pacific/Tongatapu",
    "瓦努阿图": "Pacific/Efate",
    "所罗门群岛": "Pacific/Guadalcanal",
    # ---- 非洲 ----
    "埃及": "Africa/Cairo",
    "摩洛哥": "Africa/Casablanca",
    "阿尔及利亚": "Africa/Algiers",
    "突尼斯": "Africa/Tunis",
    "利比亚": "Africa/Tripoli",
    "苏丹": "Africa/Khartoum",
    "埃塞俄比亚": "Africa/Addis_Ababa",
    "肯尼亚": "Africa/Nairobi",
    "坦桑尼亚": "Africa/Dar_es_Salaam",
    "乌干达": "Africa/Kampala",
    "卢旺达": "Africa/Kigali",
    "尼日利亚": "Africa/Lagos",
    "加纳": "Africa/Accra",
    "科特迪瓦": "Africa/Abidjan",
    "塞内加尔": "Africa/Dakar",
    "喀麦隆": "Africa/Douala",
    "刚果": "Africa/Kinshasa",
    "南非": "Africa/Johannesburg",
    "纳米比亚": "Africa/Windhoek",
    "博茨瓦纳": "Africa/Gaborone",
    "津巴布韦": "Africa/Harare",
    "赞比亚": "Africa/Lusaka",
    "莫桑比克": "Africa/Maputo",
    "安哥拉": "Africa/Luanda",
    "马达加斯加": "Africa/Antananarivo",
    "毛里求斯": "Indian/Mauritius",
}

# 拿不到 tzdata 时的回落表（UTC 偏移，夏令时按当年主流取值；外贸场景够用）
_FALLBACK_H: dict[str, float] = {
    "美国": -4, "加拿大": -4, "墨西哥": -6, "巴西": -3, "阿根廷": -3, "智利": -4,
    "英国": 1, "爱尔兰": 1, "葡萄牙": 1, "法国": 2, "德国": 2, "荷兰": 2,
    "比利时": 2, "瑞士": 2, "奥地利": 2, "意大利": 2, "西班牙": 2, "丹麦": 2,
    "瑞典": 2, "挪威": 2, "芬兰": 3, "波兰": 2, "捷克": 2, "匈牙利": 2,
    "希腊": 3, "罗马尼亚": 3, "保加利亚": 3, "爱沙尼亚": 3, "拉脱维亚": 3,
    "立陶宛": 3, "乌克兰": 3, "俄罗斯": 3, "土耳其": 3,
    "日本": 9, "韩国": 9, "新加坡": 8, "马来西亚": 8, "泰国": 7, "越南": 7,
    "菲律宾": 8, "印度尼西亚": 7, "印度": 5.5, "巴基斯坦": 5, "孟加拉国": 6,
    "阿联酋": 4, "沙特阿拉伯": 3, "以色列": 3, "伊朗": 3.5, "卡塔尔": 3,
    "科威特": 3, "埃及": 3, "南非": 2, "尼日利亚": 1, "肯尼亚": 3,
    "摩洛哥": 1, "澳大利亚": 11, "新西兰": 13, "中国": 8, "中国香港": 8,
    "中国台湾": 8, "中国澳门": 8,
}

_ZI = None
try:  # zoneinfo 需要 tzdata；venv 里有，别的解释器可能没有
    from zoneinfo import ZoneInfo

    _ = ZoneInfo("Asia/Shanghai")
    _ZI = ZoneInfo
except Exception:
    _ZI = None


def zone_of(country: str) -> str | None:
    """国家中文名 → IANA 时区名。认不出返回 None。"""
    if not country:
        return None
    c = country.strip()
    if c in _TZ:
        return _TZ[c]
    # 去掉常见后缀再试（客户表常写成「美国广交会客户」「德国\\n开发客户」）
    for cut in ("广交会", "客户", "开发", "（", "("):
        i = c.find(cut)
        if i > 0:
            c2 = c[:i].strip()
            if c2 in _TZ:
                return _TZ[c2]
    return None


def offset_hours(country: str, at: datetime | None = None) -> float | None:
    """该国在 at 时刻相对 UTC 的偏移（小时）。认不出/无 tzdata 回落固定表。"""
    z = zone_of(country)
    at = at or datetime.now(CN)
    if z and _ZI is not None:
        try:
            off = at.astimezone(_ZI(z)).utcoffset()
            if off is not None:
                return off.total_seconds() / 3600.0
        except Exception:
            pass
    # 回落
    c = (country or "").strip()
    if c in _FALLBACK_H:
        return _FALLBACK_H[c]
    for cut in ("广交会", "客户", "开发", "（", "("):
        i = c.find(cut)
        if i > 0 and c[:i].strip() in _FALLBACK_H:
            return _FALLBACK_H[c[:i].strip()]
    return None


def cn_to_local(cn_when: datetime, country: str) -> datetime | None:
    """把「北京时间」换算成「该国当地时间」。认不出国家返回 None。"""
    oh = offset_hours(country, cn_when)
    if oh is None:
        return None
    utc = cn_when.astimezone(timezone.utc)
    return utc.astimezone(timezone(timedelta(hours=oh)))


def fmt(dt: datetime) -> str:
    return dt.strftime("%Y/%m/%d %H:%M")


def parse(s: str) -> datetime | None:
    s = (s or "").replace("-", "/").replace("：", ":").strip()
    for f in ("%Y/%m/%d %H:%M", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            continue
    return None


def preview(cn_when: str, countries: list[str], limit: int = 6) -> list[str]:
    """给界面/确认框用的换算预览：['美国 当地 10/05 21:00', ...]。"""
    base = parse(cn_when)
    if base is None:
        return []
    out, seen = [], set()
    for c in countries:
        c = (c or "").strip()
        if not c or c in seen:
            continue
        seen.add(c)
        loc = cn_to_local(base, c)
        if loc is None:
            out.append(f"{c} 当地 ——（时区未收录，不换算）")
        else:
            d = ""
            # base 是 naive（北京时间），loc 是 aware → 去掉 tz 再比大小
            if loc.date() != base.date():
                d = "（前一天）" if loc.replace(tzinfo=None) < base else "（次日）"
            out.append(f"{c} 当地 {fmt(loc)}{d}")
        if len(out) >= limit:
            break
    return out
