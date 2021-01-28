from urllib.request import urlopen
from zipfile import ZipFile
from io import BytesIO
import webbrowser
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import pandas as pd

api_key = "6fd643eda53c476478c70ae1661d222b3fd1264d"
symbol = "028050" # 삼성엔지니어
standard_year = "2019"

# 종목코드로 고유번호를 찾는 함수(CORPCODE.xml이 같은 경로 안에 있어야 함)
def find_corp_code(symbol):
    # Parsing CORPCODE.xml
    parse_corp = ET.parse('CORPCODE.xml')
    corp_root = parse_corp.getroot()
    for country in corp_root.iter("list"):
        if country.findtext("stock_code") == symbol:
            return country.findtext("corp_code")

corp_code = find_corp_code(symbol)

# 고유번호 이용한 사업보고서 검색
url = "https://opendart.fss.or.kr/api/list.xml?crtfc_key="+api_key+"&pblntf_ty=A&pblntf_detail_ty=A001&last_reprt_at=Y&bgn_de="+standard_year+"0101&corp_code="+corp_code
searchXML = urlopen(url)
searchResult = searchXML.read()
xmlSoup = BeautifulSoup(searchResult,'html.parser')
#print(xmlSoup)

# 불러온 정보를 list 단위로 나누어 DF에 저장
resultDF = pd.DataFrame()
reportList = xmlSoup.findAll("list")
for t in reportList:
    temp = pd.DataFrame( ([[t.corp_cls.string, t.corp_name.string, t.corp_code.string, t.report_nm.string,
                            t.rcept_no.string, t.flr_nm.string, t.rcept_dt.string, t.rm.string]]),
                        columns=["corp_cls", "corp_name", "corp_code", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"])
    resultDF = pd.concat([resultDF, temp])

print(resultDF)

# TODO: report_nm이 '사업보고서 (기준연도.12)'인 보고서 접수번호 찾기
