#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂车帝奇瑞经销商点检 — 预留接口（尚未实现）

接入步骤（与易车保持同构，预计工作量集中在解析层）：
  1. 新建 src/dongchedi_detail.py，实现
       fetch_dealer_dongchedi(dealer: dict, features: dict, screenshot_dir: str) -> dict
     返回结构与 src/yiche_detail.fetch_dealer_yiche 完全一致：
       {province, city, dealer_id, name, pricing[], article_date, screenshot_path, error}
  2. 本文件参照 yiche_task.py 实现 main()（名单加载/断点/导出/核价均可直接复用）：
       - load_dealers_*：表头关键词改为「懂车帝UID / 懂车帝id」
       - export_to_xlsx / run_compare / run_zip_screenshots 共用现有实现
  3. 支持控制面板：main() 开头调用 _apply_external_config(cfg)
     （读取环境变量 WXDJ_CONFIG_JSON，写法照抄 yiche_task.py）
  4. control_panel.py 顶部 PLATFORMS["dongchedi"]["enabled"] 改为 True

注意：懂车帝同样可能存在云端 IP 封锁，需在客户本地宽带环境验证。
"""
import sys


def main():
    print("=" * 60)
    print("  懂车帝平台为预留接口，尚未实现。")
    print("  接入方法见本文件顶部注释。")
    print("=" * 60)
    sys.exit(1)


if __name__ == "__main__":
    main()
