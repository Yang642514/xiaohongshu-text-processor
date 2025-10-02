import logging
import os
from datetime import datetime


def setup_logger(log_dir: str, level: str = "INFO") -> logging.Logger:
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    logfile = os.path.join(log_dir, f"run_{timestamp}.log")

    logger = logging.getLogger("xhs_tool")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    # File handler
    fh = logging.FileHandler(logfile, encoding="utf-8")
    fh.setLevel(getattr(logging, level.upper(), logging.INFO))
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(getattr(logging, level.upper(), logging.INFO))
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info("Logger initialized.")
    return logger