from pydantic import BaseModel
import os
from dotenv import load_dotenv
load_dotenv(override=True, verbose=True)

class Settings(BaseModel):
    debug: bool = False
    ACCOUNT_NAME: list = ['欧莱雅黄金发膜护发直播间', '欧莱雅彩妆直播间', '欧莱雅官方旗舰店男士活力精选', '欧莱雅面膜精选', '欧莱雅染发甄选']
    START_DATE: str = '2024-01-01'
    END_DATE: str = '2024-01-31'
    BASEFLOAD: str = './uatcsv/utf-8'
    AUTHORIZATION: str = os.getenv('AUTHORIZATION')
    log_level: str = 'DEBUG'


configs = Settings()
