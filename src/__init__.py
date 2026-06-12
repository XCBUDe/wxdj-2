# ── vendor 离线依赖引导：pip 装不上时自动用随包附带的 vendor/ ──
import sys as _sys
from pathlib import Path as _Path
_vendor = _Path(__file__).resolve().parent.parent / "vendor"
if _vendor.is_dir() and str(_vendor) not in _sys.path:
    _sys.path.append(str(_vendor))   # append：优先用已正常安装的版本
del _sys, _Path, _vendor

