from sys import stdout

from loguru import logger

logger.remove()
logger.add(
    stdout,
    level="INFO",
    # encoding="utf-8",
    colorize=True,
    format="<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}",
)

logger.add(
    "log.logs",
    encoding="utf-8",
    format="<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}",
)
