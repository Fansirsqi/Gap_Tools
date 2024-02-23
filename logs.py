from sys import stdout

from loguru import logger

from config import configs

logger.remove()
logger.add(stdout, level=configs.log_level, colorize=True, format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')
logger.add('logs.log', encoding='utf-8', format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')

logger.debug(f'Test log Tools {configs.ACCOUNT_NAME}')
