from urllib.parse import quote, unquote
import os
import requests
from dotenv import load_dotenv
from config import configs

load_dotenv(override=True, verbose=True)

AUTHORIZATION = configs.AUTHORIZATION
START_DATE = configs.START_DATE
END_DATE = configs.END_DATE
ACCOUNT_NAME = os.getenv('ACCOUNT_NAME')


def got_file_name(encoded_url):
    """从响应头获取文件名"""
    return unquote(encoded_url)


_headers = {
    'Accept': '*/*',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Origin': 'https://dyaccountmgt.platform-loreal.cn',
    'Pragma': 'no-cache',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-site',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
    'authorization': AUTHORIZATION,
    'sec-ch-ua': '"Not A(Brand";v="99", "Microsoft Edge";v="121", "Chromium";v="121"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}

dy_id_hasmap = {
    '欧莱雅黄金发膜护发直播间': '120',
    '欧莱雅彩妆直播间': '121',
    '欧莱雅官方旗舰店男士活力精选': '122',
    '欧莱雅面膜精选': '123',
    '欧莱雅染发甄选': '124',
}

_params = {
    'dy_id': dy_id_hasmap[ACCOUNT_NAME],
}

rb_data = {
    'group_id': '5',
    'temp[0]': '108',
    'temp[1]': '109',
    'temp[2]': '276',
    'temp[3]': '275',
    'temp[4]': '110',
    'temp[5]': '111',
    'temp[6]': '349',
    'temp[7]': '112',
    'temp[8]': '350',
    'temp[9]': '363',
    'temp[10]': '113',
    'temp[11]': '114',
    'temp[12]': '115',
    'temp[13]': '116',
    'temp[14]': '117',
    'temp[15]': '118',
    'temp[16]': '119',
    'temp[17]': '120',
    'temp[18]': '245',
    'temp[19]': '246',
    'temp[20]': '2654',
    'temp[21]': '121',
    'temp[22]': '122',
    'temp[23]': '123',
    'temp[24]': '124',
    'temp[25]': '125',
    'temp[26]': '303',
    'temp[27]': '304',
    'temp[28]': '305',
    'temp[29]': '126',
    'temp[30]': '127',
    'temp[31]': '128',
    'temp[32]': '351',
    'temp[33]': '129',
    'temp[34]': '352',
    'temp[35]': '130',
    'temp[36]': '353',
    'temp[37]': '131',
    'temp[38]': '132',
    'temp[39]': '133',
    'temp[40]': '134',
    'temp[41]': '135',
    'temp[42]': '136',
    'temp[43]': '241',
    'temp[44]': '242',
    'temp[45]': '243',
    'temp[46]': '244',
    'temp[47]': '247',
    'temp[48]': '248',
    'start_time': START_DATE,
    'end_time': END_DATE,
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

ll_data1 = {
    'group_id': '1',
    'temp[0]': '1',
    'temp[1]': '2',
    'temp[2]': '3',
    'temp[3]': '22',
    'temp[4]': '23',
    'temp[5]': '41',
    'temp[6]': '42',
    'temp[7]': '43',
    'temp[8]': '81',
    'temp[9]': '82',
    'temp[10]': '83',
    'temp[11]': '5',
    'temp[12]': '6',
    'temp[13]': '7',
    'temp[14]': '8',
    'temp[15]': '9',
    'temp[16]': '10',
    'temp[17]': '11',
    'temp[18]': '12',
    'temp[19]': '13',
    'temp[20]': '14',
    'temp[21]': '15',
    'temp[22]': '16',
    'temp[23]': '17',
    'temp[24]': '18',
    'temp[25]': '19',
    'temp[26]': '20',
    'temp[27]': '21',
    'temp[28]': '24',
    'temp[29]': '25',
    'temp[30]': '26',
    'temp[31]': '27',
    'temp[32]': '28',
    'temp[33]': '29',
    'temp[34]': '30',
    'temp[35]': '31',
    'temp[36]': '32',
    'temp[37]': '33',
    'temp[38]': '34',
    'temp[39]': '35',
    'temp[40]': '36',
    'temp[41]': '37',
    'temp[42]': '38',
    'temp[43]': '39',
    'temp[44]': '40',
    'temp[45]': '249',
    'temp[46]': '250',
    'temp[47]': '251',
    'temp[48]': '252',
    'temp[49]': '253',
    'temp[50]': '254',
    'temp[51]': '255',
    'temp[52]': '256',
    'temp[53]': '257',
    'temp[54]': '258',
    'temp[55]': '259',
    'temp[56]': '260',
    'temp[57]': '261',
    'temp[58]': '262',
    'temp[59]': '264',
    'temp[60]': '265',
    'temp[61]': '266',
    'temp[62]': '44',
    'temp[63]': '45',
    'temp[64]': '46',
    'temp[65]': '47',
    'temp[66]': '48',
    'temp[67]': '49',
    'temp[68]': '50',
    'temp[69]': '51',
    'temp[70]': '52',
    'temp[71]': '53',
    'temp[72]': '54',
    'temp[73]': '55',
    'temp[74]': '56',
    'temp[75]': '57',
    'temp[76]': '58',
    'temp[77]': '59',
    'temp[78]': '60',
    'temp[79]': '84',
    'temp[80]': '85',
    'temp[81]': '86',
    'temp[82]': '87',
    'temp[83]': '88',
    'temp[84]': '89',
    'temp[85]': '90',
    'temp[86]': '91',
    'temp[87]': '92',
    'temp[88]': '93',
    'temp[89]': '94',
    'temp[90]': '95',
    'temp[91]': '96',
    'temp[92]': '97',
    'temp[93]': '98',
    'temp[94]': '99',
    'temp[95]': '100',
    'start_time': START_DATE,
    'end_time': END_DATE,
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

ll_data2 = {
    'group_id': '3',
    'temp[0]': '101',
    'temp[1]': '102',
    'temp[2]': '103',
    'temp[3]': '360',
    'temp[4]': '104',
    'temp[5]': '361',
    'temp[6]': '105',
    'temp[7]': '362',
    'temp[8]': '106',
    'start_time': START_DATE,
    'end_time': END_DATE,
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

rq_data = {
    'group_id': '4',
    'temp[0]': '138',
    'temp[1]': '139',
    'temp[2]': '140',
    'temp[3]': '141',
    'temp[4]': '142',
    'temp[5]': '143',
    'temp[6]': '144',
    'temp[7]': '145',
    'temp[8]': '146',
    'temp[9]': '147',
    'temp[10]': '148',
    'temp[11]': '149',
    'temp[12]': '150',
    'temp[13]': '151',
    'temp[14]': '152',
    'temp[15]': '153',
    'temp[16]': '154',
    'temp[17]': '155',
    'temp[18]': '156',
    'temp[19]': '159',
    'temp[20]': '160',
    'temp[21]': '161',
    'temp[22]': '162',
    'temp[23]': '163',
    'temp[24]': '164',
    'temp[25]': '167',
    'temp[26]': '168',
    'temp[27]': '169',
    'temp[28]': '170',
    'temp[29]': '171',
    'temp[30]': '172',
    'temp[31]': '173',
    'start_time': START_DATE,
    'end_time': END_DATE,
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

gl_data12 = {
    'group_id': '20',
    'temp[0]': '378',
    'temp[1]': '2655',
    'temp[2]': '379',
    'temp[3]': '380',
    'temp[4]': '381',
    'temp[5]': '382',
    'temp[6]': '383',
    'temp[7]': '384',
    'temp[8]': '385',
    'temp[9]': '386',
    'temp[10]': '387',
    'temp[11]': '388',
    'temp[12]': '389',
    'temp[13]': '390',
    'temp[14]': '391',
    'temp[15]': '392',
    'temp[16]': '393',
    'temp[17]': '394',
    'temp[18]': '395',
    'temp[19]': '396',
    'temp[20]': '397',
    'temp[21]': '398',
    'start_time': '2024-01-01',
    'end_time': '2024-01-31',
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

gl_data3 = {
    'group_id': '25',
    'temp[0]': '299',
    'temp[1]': '300',
    'temp[2]': '301',
    'temp[3]': '302',
    'start_time': '2024-01-01',
    'end_time': '2024-01-31',
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

gl_data4 = {
    'group_id': '21',
    'temp[0]': '364',
    'temp[1]': '365',
    'temp[2]': '366',
    'temp[3]': '367',
    'temp[4]': '368',
    'temp[5]': '369',
    'temp[6]': '370',
    'temp[7]': '372',
    'temp[8]': '373',
    'temp[9]': '374',
    'temp[10]': '375',
    'temp[11]': '376',
    'temp[12]': '377',
    'start_time': '2024-01-01',
    'end_time': '2024-01-31',
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

gl_data5 = {
    'group_id': '22',
    'temp[0]': '321',
    'temp[1]': '322',
    'temp[2]': '323',
    'temp[3]': '324',
    'temp[4]': '328',
    'temp[5]': '329',
    'temp[6]': '330',
    'temp[7]': '331',
    'temp[8]': '332',
    'temp[9]': '333',
    'temp[10]': '334',
    'temp[11]': '335',
    'temp[12]': '336',
    'temp[13]': '337',
    'temp[14]': '325',
    'temp[15]': '326',
    'temp[16]': '327',
    'temp[17]': '356',
    'temp[18]': '357',
    'start_time': '2024-01-01',
    'end_time': '2024-01-31',
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}

gl_data67 = {
    'group_id': '23',
    'temp[0]': '338',
    'temp[1]': '339',
    'temp[2]': '340',
    'temp[3]': '355',
    'temp[4]': '341',
    'temp[5]': '342',
    'temp[6]': '343',
    'temp[7]': '344',
    'temp[8]': '345',
    'temp[9]': '346',
    'temp[10]': '347',
    'temp[11]': '348',
    'start_time': '2024-01-01',
    'end_time': '2024-01-31',
    'is_save': '0',
    'is_mtd': '0',
    'is_summary': '0',
}


hosts = {
    'prod': 'https://dyaccountmgt-api.platform-loreal.cn',
    'uat': 'https://dyaccountmgt-uat-api.platform-loreal.cn',
}


def down(data: dict, host: str):
    response = requests.post(f'{host}/bo/template/export', params=_params, headers=_headers, data=data, stream=True)
    if response.status_code == 200:
        content_disposition = response.headers.get('Content-Disposition')
        if content_disposition:
            _file_name = content_disposition.split('filename=')[1].strip('"')
            file_name = got_file_name(_file_name) or 'expoet'
            if not os.path.exists(ACCOUNT_NAME):
                os.mkdir(ACCOUNT_NAME)
            down_file = f'./{ACCOUNT_NAME}/{file_name}'
            with open(down_file, 'wb') as f:
                f.write(response.content)
            print(f'下载 {file_name} OK！')
            return down_file
        else:
            print('没有找到文件名-不执行下载',response.headers)
    else:
        print(f'请求失败，状态码：{response.status_code}')


if __name__ == '__main__':
    print(ACCOUNT_NAME)
    datas = [rb_data, ll_data1, ll_data2, rq_data, gl_data12, gl_data3, gl_data4, gl_data5, gl_data67]
    for data in datas:
        down(data=data, host=hosts['uat'])
