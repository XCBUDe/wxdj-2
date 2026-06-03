# wxdj-2 — 汽车之家奇瑞经销商点检

## 快速上手（新 agent 先看这里）

```bash
pip install -r requirements.txt
# 编辑 task.py 顶部 TASK_CONFIG，填入名单和报价标准路径
python task.py
```

**唯一入口**：`task.py`（顶部 `TASK_CONFIG` 是唯一需要改的地方）

完整 skill 文档：`.claude/skills/dealer-inspection/SKILL.md`

## 目录结构

```
task.py                  ← 一键启动器，含 TASK_CONFIG
compare_prices.py        ← 报价对比（可单独跑）
src/
  config.py              ← 浏览器路径、URL 常量
  detail.py              ← Playwright 单店采集
  crawler.py             ← 批量编排 + 检查点
  export_xlsx.py         ← xlsx 导出（合并单元格 + 嵌图）
  dealers.py             ← 省市列表 API（名单模式不用）
output/                  ← 结果输出目录（gitignore 截图和检查点）
```

## 输入格式

| 文件 | 要求 |
|------|------|
| 经销商名单 `.xlsx` | Sheet1，第4行起；B列=简称，C列=uid |
| 报价标准 `.xlsx` | 「网销平台报价建议表」sheet；E=市场价，F=媒体车型名，H=现金优惠 |

## 输出

| 文件 | 内容 |
|------|------|
| `output/result.xlsx` | 全量采集结果（嵌入头图截图） |
| `output/result_checked.xlsx` | 报价对比着色版（红=价格不符，黄=标准无此车型） |
| `output/头图截图.zip` | 截图按经销商简称命名 |

## 环境要求

- Python 3.11+
- Playwright Chromium：`/opt/pw-browsers/chromium-1194/chrome-linux/chrome`
- 网络：需访问 `dealer.autohome.com.cn`
