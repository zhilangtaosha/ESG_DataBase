from urllib.request import urlopen
from zipfile import ZipFile
from io import BytesIO
import webbrowser
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import pandas as pd
import dart_fss as dart

api_key = "6fd643eda53c476478c70ae1661d222b3fd1264d"
symbol = "028050"       # 삼성엔지니어
standard_year = "2019"  # 사업보고서 연도

# # 종목코드로 고유번호를 찾는 함수(CORPCODE.xml이 같은 경로 안에 있어야 함)
# def find_corp_code(symbol):
#     # Parsing CORPCODE.xml
#     parse_corp = ET.parse('CORPCODE.xml')
#     corp_root = parse_corp.getroot()
#     for country in corp_root.iter("list"):
#         if country.findtext("stock_code") == symbol:
#             return country.findtext("corp_code")
#
# corp_code = find_corp_code(symbol)
#
# # 고유번호 이용한 사업보고서 검색
# url = "https://opendart.fss.or.kr/api/list.xml?crtfc_key="+api_key+"&pblntf_ty=A&pblntf_detail_ty=A001&last_reprt_at=Y&bgn_de="+standard_year+"0101&corp_code="+corp_code
# searchXML = urlopen(url)
# searchResult = searchXML.read()
# xmlSoup = BeautifulSoup(searchResult,'html.parser')
# #print(xmlSoup)
#
# # 불러온 정보를 list 단위로 나누어 DF에 저장
# resultDF = pd.DataFrame()
# reportList = xmlSoup.findAll("list")
# for t in reportList:
#     temp = pd.DataFrame( ([[t.corp_cls.string, t.corp_name.string, t.corp_code.string, t.report_nm.string,
#                             t.rcept_no.string, t.flr_nm.string, t.rcept_dt.string, t.rm.string]]),
#                         columns=["corp_cls", "corp_name", "corp_code", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"])
#     resultDF = pd.concat([resultDF, temp])
#
# print(resultDF)
#
# # report_nm이 '사업보고서 (기준연도.12)'인 보고서 접수번호 찾기
# rcept_no = resultDF[resultDF['report_nm']=='사업보고서 ('+standard_year+'.12)']['rcept_no']
# print(rcept_no)

# A001(사업보고서) 검색 결과 List를 DF로 변환
dart.set_api_key(api_key=api_key)
reports = dart.search(corp_code=symbol,bgn_de=standard_year+"0101",last_reprt_at="Y",pblntf_detail_ty="A001")
resultDF = pd.DataFrame()
for t in reports:
    temp = pd.DataFrame( ([[t.corp_cls, t.corp_name, t.corp_code, t.report_nm, t.rcept_no, t.flr_nm, t.rcept_dt, t.rm]]),
                        columns=["corp_cls", "corp_name", "corp_code", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"])
    resultDF = pd.concat([resultDF, temp])

# report_nm이 '사업보고서 (기준연도.12)'인 보고서 접수번호 찾기(rcept_no)
rcept_no= resultDF[resultDF['report_nm']=='사업보고서 ('+standard_year+'.12)']['rcept_no'].values
report_url = "http://dart.fss.or.kr/dsaf001/main.do?rcpNo="+''.join(rcept_no)    # series이므로 string으로 변환
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
webbrowser.open(url_final)
html_final = urlopen(url_final)
bsobjt_final = BeautifulSoup(html_final,"html.parser")
