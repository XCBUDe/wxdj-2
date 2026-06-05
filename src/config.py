"""
全局配置：品牌ID、省市映射、功能默认值、浏览器路径
"""

# 奇瑞品牌 ID（autohome内部ID，如有变化可在此修改）
CHERY_BRAND_ID = 17

# 无头浏览器启动参数（不指定 executable_path，让 Playwright 用自己安装的 Chromium）
CHROMIUM_PATH = None
CHROMIUM_ARGS = [
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--ignore-certificate-errors",
    "--disable-blink-features=AutomationControlled",
]

# 请求通用 Headers（requests 用）
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://dealer.autohome.com.cn/",
}

# 请求间隔（秒），防触发反爬
REQUEST_DELAY = 1.5

# 视口尺寸（影响截图宽度）
VIEWPORT = {"width": 1440, "height": 900}

# 中文字体（样本占位截图用，避免中文乱码）
CJK_FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"

# 截图区域：从页面顶部裁剪到 "查看更多促销" 按钮下方（约 620px）
# 若 SCREENSHOT_SELECTOR 对应元素存在，优先按元素截图
SCREENSHOT_SELECTOR = ".dealer-main-hd, .detail-top, .dealer-banner-wrap, .index-main"
SCREENSHOT_CLIP_HEIGHT = 640   # fallback 固定高度 px

# 省份配置（provinceId 为 autohome 内部值，city list 由 API 动态拉取）
# autohome 省份ID可从 https://dealer.autohome.com.cn/ 抓包确认
PROVINCE_CONFIG = {
    "广东": {"id": 19},
    "福建": {"id": 7},
    "海南": {"id": 9},
}

# 城市列表 API（GET，返回 JSON）
CITY_LIST_API = "https://dealer.autohome.com.cn/pc/area/getcitylist?provinceId={province_id}"

# 经销商列表 API
DEALER_LIST_API = (
    "https://dealer.autohome.com.cn/pc/dealer/getdealerlistbyfilter"
    "?provinceId={province_id}&cityId={city_id}&brandId={brand_id}&pageIndex={page}&pageSize=20"
)

# 经销商首页 URL 模板
DEALER_PAGE_URL = "https://dealer.autohome.com.cn/{dealer_id}"

# 功能默认开关（可通过 run.py 参数覆盖）
DEFAULT_FEATURES = {
    "pricing": True,        # 车型报价
    "screenshot": True,     # 头图截图
    "article_date": True,   # 最新软文日期
}

# 输出目录
OUTPUT_DIR = "output"
SCREENSHOT_DIR = "output/screenshots"
