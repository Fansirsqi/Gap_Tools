from pydantic import BaseModel


class Settings(BaseModel):
    debug: bool = False
    ACCOUNT_NAME: list = ['欧莱雅黄金发膜护发直播间', '欧莱雅彩妆直播间', '欧莱雅官方旗舰店男士活力精选', '欧莱雅面膜精选', '欧莱雅染发甄选']
    START_DATE: str = '2024-01-01'
    END_DATE: str = '2024-01-31'
    BASEFLOAD: str = './uatcsv/utf-8'
    AUTHORIZATION: str = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJrZXkiOjE2OSwibmFtZSI6ImRvdXlpbl9zY3JtIiwidGltZXN0YW1wIjoxNzA4NTc2NTM1LCJ1bmlxdWUiOiI2ZGYzZDRlYzZkZjE4Y2EzMmYxNzdiYTdjZDg5YTg2NSIsIm5vbmNlIjoiZWU5ZmQ2YzVlNTg0MWFlY2E4M2YxNmJiNDQ1YTEwYjMwYjFhOWIyNTlmNzExNjk5YzZiYTM5MThhMTIzYzFmZiJ9.zvy_aDJj7v2yvuEzlYDWL7F18MD2G2kh0b-6dfU3K0k'
    log_level: str = 'DEBUG'


configs = Settings()
