from urllib.request import urlopen
from zipfile import ZipFile
from io import BytesIO
import webbrowser
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import dart_fss as dart
import requests

input_path = "G:\\내 드라이브\\icin\\인증개선_icin_기업명_20210106.xlsx"
symbol_df = pd.read_excel(input_path, usecols="C", sheet_name=0,nrows=10,dtype=str) # string 형태로 저장해야 맨 앞의 0이 포함됨

api_key = "6fd643eda53c476478c70ae1661d222b3fd1264d"
standard_year = "2019"  # 사업보고서 연도


# A001(사업보고서) 검색 결과 List를 DF로 변환
dart.set_api_key(api_key=api_key)
for index in symbol_df.index:
    symbol = symbol_df['Symbol'][index]
    print(symbol)
    reports = dart.search(corp_code=symbol,bgn_de=standard_year+"0101",last_reprt_at="Y",pblntf_detail_ty="A001")
    resultDF = pd.DataFrame()
    for t in reports:
        temp = pd.DataFrame( ([[t.corp_cls, t.corp_name, t.corp_code, t.report_nm, t.rcept_no, t.flr_nm, t.rcept_dt, t.rm]]),
                            columns=["corp_cls", "corp_name", "corp_code", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"])
        resultDF = pd.concat([resultDF, temp])

# report_nm이 '사업보고서 (기준연도.12)'인 보고서 접수번호 찾기(rcept_no)
    print(resultDF['report_nm'])
    rcept_no= resultDF[resultDF['report_nm'].str.contains(standard_year)]['rcept_no'].values
    report_url = "http://dart.fss.or.kr/dsaf001/main.do?rcpNo="+''.join(rcept_no)    # series이므로 string으로 변환
    print(report_url)
    html = urlopen(report_url)
    bsobjt = BeautifulSoup(html,"html.parser")
    body = str(bsobjt.find('head'))
    body = body.split('임원의 보수')[1]
    body = body.split('cnt++')[0]
    body = body.split('viewDoc(')[1]
    body = body.split(')')[0]
    body = body.split(', ')
    body = [body[i][1:-1] for i in range(len(body))]

    url_final = 'http://dart.fss.or.kr/report/viewer.do?rcpNo=' + body[0] + '&dcmNo=' + body[1] + '&eleId=' + body[2] + '&offset=' + body[3] +'&length=' + body[4] + '&dtd=dart3.xsd'

    # 임원보수 html 가지고 오기
    #webbrowser.open(url_final)
    #html_final = urlopen(url_final)
    #bsobjt_final = BeautifulSoup(html_final,"html.parser")
    req = requests.get(url_final)
    html_final = BeautifulSoup(req.text,'html.parser')
    tables = html_final.select('table')
    #limit_table = tables[2].select('tr')
    #limit_names = limit_table[1].find('td')
    #print(limit_names)
    limit_names = []
    for tr in tables[2].select('tbody>tr'):
        tds = tr.select('td')
        limit_names.append(tds[0].text)

    print(limit_names)
    del reports, resultDF, rcept_no, report_url, html, body, url_final, html_final, tables, limit_names

# for value in enumerate(limit_table):
#     limit_names = limit_table.find('td')
# print(limit_names)
# print(tbody)
# /html/body/table[3]/tbody/tr/td[1]
# /html/body/table[3]/tbody/tr[2]/td[1]
