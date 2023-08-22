
from datetime import datetime

def do_logger(*args, **kwargs):
    current_now = datetime.now().strftime('%H:%M:%S.%f')
    log_message = current_now + ': ' + ', '.join(map(str, args)) + ', '.join(f'{k}={v}' for k, v in kwargs.items())
    with open('logger.log', mode='a+', encoding='utf-8') as log:
        log.write(log_message + '\n')

import re

class Log:
    @staticmethod
    def trans_str(_str):
        result = re.sub(r'[\(\),\'"\n\t\r]', '', str(_str))
        return result

    font_yellow = '\033[1;33m'
    font_red = '\033[1;31m'
    font_blue = '\033[1;34m'
    font_gray = '\033[1;30m'
    font_green = '\033[1;32m'
    font_purple = '\033[1;35m'
    font_cyan = '\033[1;36m'
    font_white = '\033[1;37m'
    
    bg_red = '\033[41m'
    bg_green = '\033[42m'
    bg_yellow = '\033[43m'
    bg_blue = '\033[44m'
    bg_purple = '\033[45m'
    bg_cyan = '\033[46m'
    bg_gray = '\033[47m'
    reset = '\033[0m'
    
    @staticmethod
    def warning(*context):
        message = '⚠️  [WARNING] |'
        context = Log.trans_str(context)
        print(f'{Log.font_yellow}{message}\n{context} {Log.reset}')

    @staticmethod
    def error(*context):
        message = '🔴 [ERROR]   |'
        context = Log.trans_str(context)
        print(f'{Log.font_red}{message}\n{context} {Log.reset}')

    @staticmethod
    def info(*context):
        message = '🔵 [INFO]    |'
        context = Log.trans_str(context)
        print(f'{Log.font_blue}{message}\n{context} {Log.reset}')

    @staticmethod
    def success(*context):
        message = '🟢 [SUCCESS] |'
        context = Log.trans_str(context)
        print(f'{Log.font_green}{message}\n{context} {Log.reset}')

    @staticmethod
    def debug(*context):
        message = '⚙️  [DEBUG]   |'
        context = Log.trans_str(context)
        print(f'{Log.font_gray}{message}\n{context} {Log.reset}')

    @staticmethod
    def input(context):
        message = '✍️  [INPUT]   |'
        print(f'{Log.font_white}{message}\n{context} {Log.reset}')
        data = input()
        return data