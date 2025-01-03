from loguru import logger
import os
from pathlib import Path
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 配置日志
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)

logger.add(
    os.getenv("LOG_FILE", "app.log"),
    rotation=os.getenv("LOG_ROTATION", "10 MB"),
    retention=os.getenv("LOG_RETENTION", "30 days"),
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
    enqueue=True,
    backtrace=True,
    diagnose=True
)

def get_logger():
    return logger
