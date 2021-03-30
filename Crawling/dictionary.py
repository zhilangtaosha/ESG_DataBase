from urllib.request import urlopen
import webbrowser
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import dart_fss as dart
import requests

input_path = "G:\\내 드라이브\\icin\\인증개선_icin_기업명_20210106.xlsx"
api_key = "6fd643eda53c476478c70ae1661d222b3fd1264d"
standard_year = "2019"  # 사업보고서 연도
writer = pd.ExcelWriter('G:\\내 드라이브\\DART\\보수한도테스트.xlsx', engine='xlsxwriter')

# A001(사업보고서) 검색 결과 List를 DF로 변환
# string 형태로 저장해야 맨 앞의 0이 포함됨a
symbol_df = pd.read_excel(input_path, usecols=[1,2,3], sheet_name=0, dtype=str, skiprows=19, header=None, names=['ASymbol','Symbol','기업명1'], nrows=3)
dart.set_api_key(api_key=api_key)

# 빈 dataframe 생성
payments_df = pd.DataFrame(columns=['ASymbol','Symbol','회사명','이사 구분','보수 한도(백만)','구분 전체','한도 전체','단위','확인필요'])

for index in symbol_df.index:
    print("..."+str(index+1)+" "+symbol_df['기업명1'][index]+"...\n")
    symbol = symbol_df['Symbol'][index]
    payments_df = payments_df.append({'ASymbol': symbol_df['ASymbol'][index], 'Symbol': symbol, '회사명': symbol_df['기업명1'][index]}, ignore_index=True)

    # 사업보고서(A001) 검색
    # 비상장기업 등 사업보고서 없는 경우 continue 하여 해당 기업 건너뛰기
    try:
        reports = dart.search(corp_code=symbol,bgn_de=standard_year+"0101",last_reprt_at="Y",pblntf_detail_ty="A001")
    except:
        payments_df.at[index, '이사 구분']='error'
        payments_df.at[index, '보수 한도(백만)']='error'
        continue
    resultDF = pd.DataFrame()
    for t in reports:
        temp = pd.DataFrame( ([[t.corp_cls, t.corp_name, t.corp_code, t.report_nm, t.rcept_no, t.flr_nm, t.rcept_dt, t.rm]]),
                            columns=["corp_cls", "corp_name", "corp_code", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"])
        resultDF = pd.concat([resultDF, temp])

    # 보고서 이름(report_nm)에 '기준연도'가 포함된 보고서 접수번호(rcept_no) 찾기
    rcept_no= resultDF[resultDF['report_nm'].str.contains(standard_year)]['rcept_no'].values
    report_url = "http://dart.fss.or.kr/dsaf001/main.do?rcpNo="+''.join(rcept_no)    # series이므로 string으로 변환

    # 보고서 내에서 '임원의 보수' 탭 찾기
    # 기재정정된 경우 '임원의 보수' 탭이 없음. (예: 컨버즈) 이 경우에는 continue 하여 해당 기업 건너뛰기
    html = urlopen(report_url)
    bsobjt = BeautifulSoup(html,"html.parser")
    body = str(bsobjt.find('head'))

    try:
        body = body.split('임원의 보수')[1]
    except:
        payments_df.at[index, '이사 구분']='error'
        payments_df.at[index, '보수 한도(백만)']='error'
        continue
    body = body.split('cnt++')[0]
    body = body.split('viewDoc(')[1]
    body = body.split(')')[0]
    body = body.split(', ')
    body = [body[i][1:-1] for i in range(len(body))]

    url_final = 'http://dart.fss.or.kr/report/viewer.do?rcpNo=' + body[0] + '&dcmNo=' + body[1] + '&eleId=' + body[2] + '&offset=' + body[3] +'&length=' + body[4] + '&dtd=dart3.xsd'

    # 임원보수 html 가지고 오기
    req = requests.get(url_final)
    html_final = BeautifulSoup(req.text,'html.parser')
    tables = html_final.select('table')

    # 보수 단위
    unit = ((tables[1].select('tbody>tr>td'))[1].text)[5:-1]

    # 보수 table
    limit_names = []
    limit_pays = []
    limits = {}
    for tr in tables[2].select('tbody>tr'):
        tds = tr.select('td')
        director_name = tds[0].text.replace('\n','').replace('\xa0','').replace(' ','')
        limit_names.append(director_name)     # 첫 번째 column (이사 구분)
        limits.setdefault(director_name)
        try:
            limit_pays.append(int(tds[2].text.replace('\n','').replace('\xa0','').replace(',',''))) # 세 번째 column (보수금액 integer로)
            limits.update({director_name: int(tds[2].text.replace('\n','').replace('\xa0','').replace(',',''))})
        except:
            pass

    # 단위 연산
    unit_pays = limit_pays
    if '백만' not in unit:
        unit_pays = [pay/1000. for pay in limit_pays]
        if '천' not in unit:
            unit_pays = [pay/1000. for pay in unit_pays]

    # 페이지 안의 table 순서나 개수가 달라서 crawling이 안되는 경우가 있다. (예: 만호제강) 이 경우에는 'error' 입력한다.
    # 값이 여러 개일 경우 (거의 확실히) 첫 번째 값이 사내이사이므로 각 리스트의 첫 번째 값을 이사명 및 보수로 입력한다.
    try:
        payments_df.at[index, '이사 구분'] = limit_names[0]
        payments_df.at[index, '보수 한도(백만)'] = unit_pays[0]
        payments_df.at[index, '구분 전체'] = '//'.join(limit_names)
        payments_df.at[index, '한도 전체'] = '//'.join(map(str, limit_pays))
        payments_df.at[index, '단위'] = unit
    except:
        payments_df.at[index, '이사 구분'] = 'error'
        payments_df.at[index, '보수 한도(백만)'] = 'error'

    del reports, resultDF, rcept_no, report_url, html, body, url_final, html_final, tables, limit_names, limit_pays, unit_pays, unit, limits


payments_df.insert(0, 'no.', range(1, len(payments_df) + 1))  # numbering column 추가
payments_df.to_excel(writer,index=False)
writer.close()
