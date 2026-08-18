from __future__ import annotations

import os
import sys
from typing import Literal

from .build_flags import DEPLOY_MODE_BUILD

DeployMode = Literal["dev", "pilot"]


def is_frozen_binary() -> bool:
    if getattr(sys, "frozen", False):
        return True
    # Nuitka
    if "__compiled__" in globals():
        return True
    try:
        import __main__

        if hasattr(__main__, "__compiled__"):
            return True
    except Exception:
        pass
    exe = (sys.executable or "").lower()
    if exe.endswith(".exe") and "python" not in os.path.basename(exe).lower():
        return True
    return False


def get_deploy_mode() -> DeployMode:
    """Pilot binary is always pilot; env cannot downgrade frozen builds."""
    build = (DEPLOY_MODE_BUILD or "").strip().lower()
    if build == "pilot" or is_frozen_binary():
        return "pilot"
    env = (os.environ.get("KBOND_DEPLOY_MODE") or "dev").strip().lower()
    if env == "pilot":
        return "pilot"
    return "dev"


def is_pilot() -> bool:
    return get_deploy_mode() == "pilot"


def is_dev() -> bool:
    return get_deploy_mode() == "dev"
