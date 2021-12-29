# -*- coding: utf-8 -*-
import json
import os
import time
import openpyxl
import requests
def cell2html(value):
    value = str(value)
    h = '<html><body>'
    for _ in value.split('\n'):
        h += '<p>' + _ + '</p>'
    h += '</body></html>'
    return h


file = ''
for file in os.listdir(os.getcwd()):
    if file.endswith('xlsx'):
        break
if file == '':
    print('no xlsx file found')
else:
    print('read', file)

work_book = openpyxl.open(os.path.join(os.getcwd(), file))
sheet = work_book[work_book.sheetnames[0]]

CN = sheet['C3'].value
EN = sheet['D3'].value
TW = sheet['E3'].value
KR = sheet['F3'].value
JP = sheet['G3'].value
VN = sheet['H3'].value
TH = sheet['I3'].value
ID = sheet['J3'].value
RU = sheet['K3'].value
# ini = 'D:\gonggaoconfig.ini'
# version = CN[:CN.index('版')]
ini = 'D:\\WorkSpace\\jenkins\\workspace\\player_copygonggao_tomis\\start_build\\gonggaoconfig.ini'
with open(ini, 'r', encoding='utf-8') as p:
    li = p.readlines()
dic = dict(value.replace('\n', '').split(' = ') for value in li if '=' in value)
player_version = dic['player_version']
print(player_version)
timestamp = int(1000 * time.time())
body = {'noticeType': 1, 'innerVersion': player_version, 'state': 1,
        'updateTime': timestamp, 'createTime': timestamp, 'isDelete': 0,
        'releaseNote': {
            'nox_cn': cell2html(CN),
            'default': cell2html(EN),
            'nox_en': cell2html(EN),
            'nox_tw': cell2html(TW),
            'nox_ko': cell2html(KR),
            'nox_jp': cell2html(JP),
            'nox_vn': cell2html(VN),
            'nox_th': cell2html(TH),
            'nox_id': cell2html(ID),
            'nox_ru': cell2html(RU)
        }}

r = requests.post(
    'http://10.0.8.69:10089/api/release/notice/saveOrUpdate',
    headers={'Host': '10.0.8.69:10089',
             'Connection': 'keep-alive',
             'Accept': 'application/json, text/plain, */*',
             'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
             'Content-Type': 'application/json;charset=UTF-8',
             'Origin': 'http://10.0.8.69:10089',
             'Referer': 'http://10.0.8.69:10089/new',
             'Accept-Encoding': 'gzip, deflate',
             'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7,vi;q=0.6,ko;q=0.5,ja;q=0.4,zh-CN;q=0.3,th;q=0.2,fr-FR;q=0.1,fr;q=0.1',
             'Cookie': 'JSESSIONID=A3D610DDAA59ED17C955B088996EF780'},
    data=json.dumps(body, ensure_ascii=False).encode('UTF-8')
)
print('Response status:', r.status_code)
if r.status_code == 200:
    print('Response:')
    print(r.json())
with open(ini, 'r', encoding='utf-8') as p:
    li = p.readlines()
dic = dict(value.replace('\n', '').split(' = ') for value in li if '=' in value)
rom5version = dic['rom5version']
rom7version = dic['rom7version']
rom8version = dic['rom8version']
rom9version = dic['rom9version']
_00_ = {
    '4': {
        'CN': 'ROM-V'+rom5version+'版本更新：优化安卓5(32位)相关性能',
        'EN': 'ROM-V'+rom5version+'Release Note: Android 5(32-bit) performance optimized',
        'TW': 'ROM-V'+rom5version+': 優化 安卓5(32-bit)版本相關性能',
        'KR': 'ROM-V'+rom5version+': 안드로이드 5(32비트) 호환성 개선',
        'JP': 'ROM-V'+rom5version+': Android 5(32-bit) 環境上のパフォーマンス向上しました。',
        'VN': 'ROM-V'+rom5version+': Tối ưu hóa tính năng trên giả lập Android 5 (32 bit)',
        'TH': 'ROM-V'+rom5version+': เพิ่มประสิทธิภาพ Android 5(32 บิต)',
        'ID': 'ROM-V'+rom5version+': Mengoptimalkan Kinerja Android 5(32-bit)',
        'RU': 'ROM-V'+rom5version+': Оптимизирована производительность Android 5 (32-бит)'
    },
    '5': {
        'CN': 'ROM-V'+rom7version+': 优化安卓7(32位)相关性能',
        'EN': 'ROM-V'+rom7version+': Android 7 (32-bit) performance optimized',
        'TW': 'ROM-V'+rom7version+': 1.優化 安卓7(32-bit)版本相關性能',
        'KR': 'ROM-V'+rom7version+': 안드로이드 7(32비트) 호환성 개선',
        'JP': 'ROM-V'+rom7version+': Android 7 (32-bit) 環境上のパフォーマンス向上しました。',
        'VN': 'ROM-V'+rom7version+': Tối ưu hóa tính năng trên giả lập Android 7 (32 bit)',
        'TH': 'ROM-V'+rom7version+': เพิ่มประสิทธิภาพ Android 7 (32 บิต)',
        'ID': 'ROM-V'+rom7version+': Mengoptimalkan Kinerja Android 7(32-bit)',
        'RU': 'ROM-V'+rom7version+': Оптимизирована производительность Android 7 (32-бит)'
    },
    '6': {
        'CN': 'ROM-V'+rom8version+': 优化安卓7(64位)相关性能',
        'EN': 'ROM-V'+rom8version+': Android 7(64-bit) performance optimized',
        'TW': 'ROM-V'+rom8version+': 優化 安卓7(64-bit)版本相關性能',
        'KR': 'ROM-V'+rom8version+': 안드로이드 7(64비트) 호환성 개선',
        'JP': 'ROM-V'+rom8version+': Android 7(64-bit) 環境上のパフォーマンス向上しました。',
        'VN': 'ROM-V'+rom8version+': Tối ưu hóa tính năng trên giả lập Android 7 (64 bit)',
        'TH': 'ROM-V'+rom8version+': เพิ่มประสิทธิภาพ Android 7(64 บิต)',
        'ID': 'ROM-V'+rom8version+': Mengoptimalkan Kinerja Android 7(64-bit)',
        'RU': 'ROM-V'+rom8version+': Оптимизирована производительность Android 7 (64-бит)'
    },
    '7': {
        'CN': 'ROM-V'+rom9version+': 优化安卓9(64位)相关性能',
        'EN': 'ROM-V'+rom9version+': Android 9(64-bit) performance optimized',
        'TW': 'ROM-V'+rom9version+': 1.優化 安卓9(64-bit)版本相關性能',
        'KR': 'ROM-V'+rom9version+': 안드로이드 9(64비트) 호환성 개선',
        'JP': 'ROM-V'+rom9version+': Android 9(64-bit) 環境上のパフォーマンス向上しました。',
        'VN': 'ROM-V'+rom9version+': Tối ưu hóa tính năng trên giả lập Android 9 (64 bit)',
        'TH': 'ROM-V'+rom9version+': เพิ่มประสิทธิภาพ Android 9(64 บิต)',
        'ID': 'ROM-V'+rom9version+': Mengoptimalkan Kinerja Android 9(64-bit)',
        'RU': 'ROM-V'+rom9version+': Оптимизирована производительность Android 9 (64-бит)'
    }
}

for row in ('4', '5', '6', '7'):
    CN = sheet['C' + row].value
    if not CN:
        CN = _00_[row]['CN']
    else:
        version = CN[CN.index('V'):CN.index('版')]
    EN = sheet['D' + row].value
    if not EN:
        EN = _00_[row]['EN']
    TW = sheet['E' + row].value
    if not TW:
        TW = _00_[row]['TW']
    KR = sheet['F' + row].value
    if not KR:
        KR = _00_[row]['KR']
    JP = sheet['G' + row].value
    if not JP:
        JP = _00_[row]['JP']
    VN = sheet['H' + row].value
    if not VN:
        VN = _00_[row]['VN']
    TH = sheet['I' + row].value
    if not TH:
        TH = _00_[row]['TH']
    ID = sheet['J' + row].value
    if not ID:
        ID = _00_[row]['ID']
    RU = sheet['K' + row].value
    if not RU:
        RU = _00_[row]['RU']
    print('version:', version)
    timestamp = int(1000 * time.time())
    body = {'noticeType': 2, 'innerVersion': version, 'state': 1,
            'updateTime': timestamp, 'createTime': timestamp, 'isDelete': 0}
    params = {'nox_cn': CN, 'default': EN, 'nox_en': EN, 'nox_tw': TW, 'nox_ko': KR, 'nox_jp': JP, 'nox_vn': VN,
              'nox_th': TH, 'nox_id': ID, 'nox_ru': RU}
    releaseNote = {}
    for k, v in params.items():
        releaseNote[k] = cell2html(v)
    body['releaseNote'] = releaseNote
    r = requests.post(
        'http://10.0.8.69:10089/api/release/notice/saveOrUpdate',
        headers={'Host': '10.0.8.69:10089',
                 'Connection': 'keep-alive',
                 'Accept': 'application/json, text/plain, */*',
                 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
                 'Content-Type': 'application/json;charset=UTF-8',
                 'Origin': 'http://10.0.8.69:10089',
                 'Referer': 'http://10.0.8.69:10089/new',
                 'Accept-Encoding': 'gzip, deflate',
                 'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7,vi;q=0.6,ko;q=0.5,ja;q=0.4,zh-CN;q=0.3,th;q=0.2,fr-FR;q=0.1,fr;q=0.1',
                 'Cookie': 'JSESSIONID=A3D610DDAA59ED17C955B088996EF780'},
        data=json.dumps(body, ensure_ascii=False).encode('UTF-8')
    )
    print('Response status:', r.status_code)
    if r.status_code == 200:
        print('Response:')
        print(r.json())
