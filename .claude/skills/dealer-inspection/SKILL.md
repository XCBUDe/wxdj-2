---
name: dealer-inspection
description: >
  汽车之家奇瑞经销商点检任务。给定经销商名单（Excel）和报价标准（Excel），
  批量抓取汽车之家经销商页面的店头名称、车型裸车价、最新软文日期、头图截图，
  再与报价标准对比着色输出 xlsx，并打包截图 ZIP。
  适用场景：品牌方定期核查各地经销商的网销价格执行情况。
---

# 汽车之家奇瑞经销商点检 Skill

## 概述

本 skill 执行以下完整流程：

```
名单 Excel (uid 列)
  ↓
批量 Playwright 抓取
  ├─ 店头名称（页面 title 正则）
  ├─ 车型裸车价（价格页文本解析）
  ├─ 最新软文日期（body text 正则）
  └─ 头图截图（viewport clip，PNG）
  ↓
检查点断续（JSONL）
  ↓
导出 result.xlsx（经销商合并行 + 嵌入截图）
  ↓
与报价标准对比（指导价定位 + 0容错核价）
  ├─ 裸车价不符 → 红色高亮
  └─ 标准无此车型 → 黄色高亮
  ↓
输出 result_checked.xlsx + 截图.zip
```

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
playwright install chromium 2>/dev/null || true
```

### 2. 准备输入文件

| 文件 | 说明 | 必填 |
|------|------|------|
| 经销商名单 `.xlsx` | Sheet1，第4行起；B列=简称，C列=uid（汽车之家经销商ID） | ✅ |
| 报价标准 `.xlsx` | 「网销平台报价建议表」sheet；E列=市场价格，F列=对应媒体车型，H列=现金优惠金额 | ✅（如需价格核对）|

### 3. 配置 task.py

打开 `task.py`，修改顶部 `TASK_CONFIG` 中的路径和参数即可，**无需改任何其他代码**。

```python
TASK_CONFIG = {
    "dealer_list_xlsx":   "上传的名单.xlsx",          # ← 改这里
    "standard_xlsx":      "上传的报价标准.xlsx",       # ← 改这里
    "output_dir":         "output",
    "result_xlsx":        "output/result.xlsx",
    "checked_xlsx":       "output/result_checked.xlsx",
    "screenshots_zip":    "output/头图截图.zip",
    "checkpoint":         "output/checkpoint.jsonl",
    "concurrency":        2,                            # 并发浏览器数，建议 1-3
    "features": {
        "pricing":       True,   # 采集车型报价
        "screenshot":    True,   # 采集头图截图
        "article_date":  True,   # 采集最新软文日期
    },
    "compare_prices":     True,  # 是否执行报价对比着色
    "zip_screenshots":    True,  # 是否打包截图 ZIP
}
```

### 4. 运行

```bash
python task.py
```

断点续跑（中途中断后再跑，已完成经销商自动跳过）：

```bash
python task.py          # 直接重跑，checkpoint 文件自动续跑
```

---

## 输出文件说明

| 文件 | 内容 |
|------|------|
| `output/result.xlsx` | 全量数据（每车型一行，经销商级字段合并，嵌入头图截图） |
| `output/result_checked.xlsx` | 同上 + 裸车价着色（红=不符，黄=无法核对）+ 「核对结果汇总」sheet |
| `output/头图截图.zip` | 各经销商头图 PNG，按简称重命名 |
| `output/checkpoint.jsonl` | 断点文件（每完成一家追加一行 JSON，已加入 .gitignore） |
| `output/screenshots/*.png` | 原始截图（按 dealer_id 命名） |

---

## 架构说明（供调试/扩展）

```
task.py                  ← 唯一入口，含 TASK_CONFIG
src/
  config.py              ← 浏览器路径、URL 模板、全局常量
  detail.py              ← Playwright 单店采集（名称/报价/日期/截图）
  crawler.py             ← 批量编排（ThreadPoolExecutor + 检查点）
  export_xlsx.py         ← openpyxl 导出（合并单元格 + 嵌图）
  dealers.py             ← 省市列表 API（可选，名单模式不用）
compare_prices.py        ← 报价标准对比着色（也可单独跑）
```

### 关键技术点

**报价提取**：不依赖 CSS 选择器（汽车之家用 Tailwind 动态类名），直接
`page.inner_text("body")` 后用状态机解析「车型报价」区段。

**展开所有车系**：JS 直接 click 所有含「展开」文字的 button，绕过可见性检测：
```js
document.querySelectorAll('button').forEach(btn => {
    if (btn.textContent.includes('展开')) btn.click();
});
```

**线程安全**：Playwright sync_api 不能跨线程共享，每个 worker 独立创建
`sync_playwright().start()` + `chromium.launch()`。

**报价核对匹配**：
- 车系映射表 `SERIES_MAP`（我方名 → 标准表名）
- 用**指导价（厂方市场价格）**在映射车系内唯一定位车型（比车型名更稳定）
- 同价多行时用车型名字符集 Jaccard 相似度消歧
- 裸车价 = 市场价格 − 现金优惠金额（2026.6.2 起那列 H 列）

---

## 常见问题

**Q: 某经销商截图为空或 0 车型**

Playwright 有时被反爬拦截（body < 200 字）。中断后重跑自动跳过已完成，
被拦的重试方法：降低并发到 1，或把 `REQUEST_DELAY` 调大到 3s。

**Q: 想只核价不重新采集**

```bash
python compare_prices.py        # 直接读 output/result_full.xlsx 对比
```
或修改 `task.py` 中 `TASK_CONFIG["compare_prices"] = True` 但先把
`output/result.xlsx` 换成已有结果再跑。

**Q: 想换省/市/品牌**

当前实现固定奇瑞（`CHERY_BRAND_ID = 17`）。换品牌需在
`src/config.py` 改 `CHERY_BRAND_ID`，并更新 `compare_prices.py`
的 `SERIES_MAP`。

**Q: 标准 Excel 格式变化了**

`compare_prices.py` 里 `load_standard()` 读取列：
- E 列（col 5）= 市场价格
- F 列（col 6）= 对应媒体车型
- H 列（col 8）= 现金优惠金额（2026.6.2 那列）

若标准表列顺序变了，修改这三个列号即可。

---

## 执行本 skill 的步骤（给 agent 看）

1. 确认用户已提供：经销商名单 xlsx 路径、报价标准 xlsx 路径
2. 编辑 `task.py` 顶部 `TASK_CONFIG`，填入正确路径
3. `pip install -r requirements.txt`
4. `python task.py`
5. 等待完成，把 `output/result_checked.xlsx` 和 `output/头图截图.zip` 发给用户
6. 若有标红条目，告知用户哪些经销商的哪些车型裸车价与标准不符
