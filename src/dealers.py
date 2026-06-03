"""
Step 1 — 经销商列表
输入：省份名 + 城市名（或"全部"）+ 奇瑞 brandId
输出：List[dict]  每项包含 {province, city, city_id, dealer_id, name}
"""
import time
import requests
from .config import (
    CHERY_BRAND_ID, PROVINCE_CONFIG,
    CITY_LIST_API, DEALER_LIST_API, REQUEST_HEADERS, REQUEST_DELAY
)


def _session():
    s = requests.Session()
    s.headers.update(REQUEST_HEADERS)
    return s


def get_cities(province_name: str, session=None) -> list[dict]:
    """拉取省内城市列表，返回 [{city_id, city_name}]"""
    pconf = PROVINCE_CONFIG.get(province_name)
    if not pconf:
        raise ValueError(f"未知省份: {province_name}，可选: {list(PROVINCE_CONFIG)}")
    session = session or _session()
    url = CITY_LIST_API.format(province_id=pconf["id"])
    try:
        r = session.get(url, timeout=15)
        r.raise_for_status()
        data = r.json()
        # autohome 通常返回 {"returncode":0, "result":[{"cityId":X,"cityName":"广州"},...]}
        items = data.get("result") or data.get("data") or []
        if isinstance(items, list) and items:
            # 兼容多种 key 名
            cities = []
            for it in items:
                cid = it.get("cityId") or it.get("id") or it.get("CityId")
                cname = it.get("cityName") or it.get("name") or it.get("CityName")
                if cid and cname:
                    cities.append({"city_id": cid, "city_name": cname})
            return cities
    except Exception as e:
        print(f"  [WARN] 获取城市列表失败({province_name}): {e}")
    return []


def get_dealers_in_city(province_name: str, city_id: int, city_name: str,
                         session=None) -> list[dict]:
    """拉取某城市的奇瑞经销商列表（自动翻页）"""
    pconf = PROVINCE_CONFIG[province_name]
    session = session or _session()
    dealers = []
    page = 1

    while True:
        url = DEALER_LIST_API.format(
            province_id=pconf["id"],
            city_id=city_id,
            brand_id=CHERY_BRAND_ID,
            page=page,
        )
        try:
            r = session.get(url, timeout=15)
            r.raise_for_status()
            data = r.json()
            # 解析 dealer 列表，适配多种可能的 key
            items = (
                data.get("result")
                or data.get("data")
                or data.get("dealerList")
                or []
            )
            if not items:
                break
            for it in items:
                dealer_id = (
                    it.get("dealerId") or it.get("id") or it.get("DealerId")
                )
                name = (
                    it.get("dealerName") or it.get("name") or it.get("DealerName")
                )
                if dealer_id:
                    dealers.append({
                        "province": province_name,
                        "city": city_name,
                        "city_id": city_id,
                        "dealer_id": str(dealer_id),
                        "name": name or "",
                    })
            # 翻页：若返回数量 < 20，则没有下一页
            if len(items) < 20:
                break
            page += 1
            time.sleep(REQUEST_DELAY)
        except Exception as e:
            print(f"  [WARN] 获取经销商列表失败 {city_name} p{page}: {e}")
            break

    return dealers


def get_all_dealers(province: str = "全部", city: str = "全部") -> list[dict]:
    """
    主入口：按省/市过滤，返回全部匹配经销商列表。
    province: "广东"/"福建"/"海南"/"全部"
    city:     城市名称（如"东莞"）/"全部"
    """
    session = _session()
    provinces = (
        list(PROVINCE_CONFIG.keys()) if province == "全部" else [province]
    )
    all_dealers = []

    for pname in provinces:
        print(f"[INFO] 拉取城市列表: {pname}")
        cities = get_cities(pname, session)
        if not cities:
            print(f"  [WARN] {pname} 无城市数据，跳过")
            continue

        if city != "全部":
            cities = [c for c in cities if c["city_name"] == city]
            if not cities:
                print(f"  [WARN] {pname} 中未找到城市: {city}")
                continue

        for c in cities:
            print(f"  [INFO] 拉取经销商: {pname}-{c['city_name']}")
            dealers = get_dealers_in_city(pname, c["city_id"], c["city_name"], session)
            print(f"    -> {len(dealers)} 家")
            all_dealers.extend(dealers)
            time.sleep(REQUEST_DELAY)

    return all_dealers


# ── 样本数据（--sample 模式 / 单元测试用）──────────────────────────────
# 仅保留用户提供的真实经销商（清远冠荣 = 2045891）
SAMPLE_DEALERS = [
    {"province": "广东", "city": "清远", "city_id": 37,
     "dealer_id": "2045891", "name": "奇瑞清远冠荣体验中心"},
]
