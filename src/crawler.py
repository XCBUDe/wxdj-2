"""
Step 3 — 批量编排（支持断点续跑、并发）
"""
import json
import time
import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from .config import REQUEST_DELAY
from .dealers import get_all_dealers, SAMPLE_DEALERS
from .detail import fetch_dealer_detail, SAMPLE_DETAILS


def _load_checkpoint(path: str) -> set[str]:
    """读取已完成的 dealer_id 集合"""
    done = set()
    if Path(path).exists():
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        done.add(json.loads(line)["dealer_id"])
                    except Exception:
                        pass
    return done


def _append_checkpoint(path: str, record: dict):
    """追加一条记录到断点文件"""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def _load_all_from_checkpoint(path: str) -> list[dict]:
    """读取断点文件中全部记录"""
    records = []
    if Path(path).exists():
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        records.append(json.loads(line))
                    except Exception:
                        pass
    return records


def _try_real_screenshot(dealer_id: str) -> str:
    """
    用 Playwright 联网截取经销商首页头图区域。
    网络可达返回图片路径；被拦/失败返回空字符串（不造假图）。
    """
    from .detail import fetch_dealer_detail
    dealer = {"province": "广东", "city": "清远", "dealer_id": dealer_id,
              "name": "奇瑞清远冠荣体验中心"}
    res = fetch_dealer_detail(
        dealer,
        features={"pricing": False, "screenshot": True, "article_date": False},
    )
    path = res.get("screenshot_path", "")
    # 校验：文件存在且非空白（>3KB 视为有效页面截图）
    if path and Path(path).exists() and Path(path).stat().st_size > 3000:
        return path
    return ""


def run_crawl(
    province: str = "全部",
    city: str = "全部",
    features: dict = None,
    output: str = "output/result.xlsx",
    test_dealer_id: str = None,
    concurrency: int = 2,
    checkpoint_file: str = "output/checkpoint.jsonl",
    sample_mode: bool = False,
):
    """
    主采集入口。
    sample_mode=True 时使用内置样本数据，不联网。
    test_dealer_id 指定时只测单店。
    """
    from .export_xlsx import export_to_xlsx

    features = features or {"pricing": True, "screenshot": True, "article_date": True}
    Path(output).parent.mkdir(parents=True, exist_ok=True)

    # ── 样本模式 ──────────────────────────────────────────────────────────
    # 报价取自 data/standards.json 真实数据；截图字段若开启 screenshot，
    # 会尝试用 Playwright 联网截真实头图（网络可达时成功，被拦时留空并提示）。
    if sample_mode:
        print("[SAMPLE] 报价=standards.json 真实数据")
        records = []
        dealers = SAMPLE_DEALERS
        if test_dealer_id:
            dealers = [d for d in dealers if d["dealer_id"] == test_dealer_id]
        for d in dealers:
            detail = dict(SAMPLE_DETAILS.get(d["dealer_id"], {}))
            if not detail:
                detail = {**d, "pricing": [], "article_date": "", "error": "无样本"}
            detail["screenshot_path"] = ""
            if not features.get("pricing"):
                detail["pricing"] = []
            if not features.get("article_date"):
                detail["article_date"] = ""
            # 尝试联网截真实头图
            if features.get("screenshot"):
                shot = _try_real_screenshot(d["dealer_id"])
                detail["screenshot_path"] = shot
                print(f"  [截图] {'成功 ' + shot if shot else '失败(网络不可达，留空)'}")
            records.append(detail)
            print(f"  [OK] {d['dealer_id']} {detail.get('name','')} / {len(detail.get('pricing') or [])}车型")
        export_to_xlsx(records, output, features)
        print(f"\n[DONE] 已输出: {output}")
        return records

    # ── 真实采集模式 ──────────────────────────────────────────────────────
    # 1. 获取经销商列表
    if test_dealer_id:
        dealers = [{"province": province if province != "全部" else "未知",
                    "city": city if city != "全部" else "未知",
                    "city_id": 0,
                    "dealer_id": test_dealer_id,
                    "name": ""}]
    else:
        print(f"[INFO] 拉取经销商列表: 省={province} 市={city}")
        dealers = get_all_dealers(province, city)
        print(f"[INFO] 共 {len(dealers)} 家经销商")

    if not dealers:
        print("[WARN] 未找到经销商，退出")
        return []

    # 2. 断点续跑
    done_ids = _load_checkpoint(checkpoint_file)
    todo = [d for d in dealers if d["dealer_id"] not in done_ids]
    print(f"[INFO] 待采集: {len(todo)} 家（已完成: {len(done_ids)}）")

    # 3. 并发采集（共享浏览器实例以减少启动开销）
    records_from_checkpoint = _load_all_from_checkpoint(checkpoint_file)

    def _work_single(dealer: dict) -> dict:
        # Each worker thread creates its own Playwright + browser instance.
        # Playwright sync_api uses greenlets internally and cannot be shared
        # across threads.
        result = fetch_dealer_detail(dealer, features, browser=None)
        _append_checkpoint(checkpoint_file, result)
        time.sleep(REQUEST_DELAY)
        return result

    new_records = []
    try:
        with ThreadPoolExecutor(max_workers=concurrency) as pool:
            futures = {
                pool.submit(_work_single, dealer): dealer
                for dealer in todo
            }
            for i, fut in enumerate(as_completed(futures), 1):
                dealer = futures[fut]
                try:
                    rec = fut.result()
                    new_records.append(rec)
                    status = "OK" if not rec.get("error") else f"ERR:{rec['error'][:40]}"
                    print(f"  [{i}/{len(todo)}] {dealer['dealer_id']} {dealer.get('name','')} -> {status}")
                except Exception as e:
                    print(f"  [{i}/{len(todo)}] {dealer['dealer_id']} FATAL: {e}")
    except Exception as e:
        print(f"[FATAL] 并发采集异常: {e}")

    # 4. 合并并导出
    all_records = records_from_checkpoint + new_records
    export_to_xlsx(all_records, output, features)
    print(f"\n[DONE] 共 {len(all_records)} 条记录，已输出: {output}")
    return all_records
