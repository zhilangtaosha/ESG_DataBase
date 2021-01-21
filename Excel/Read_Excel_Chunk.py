'''
pd.read_excel()은 더 이상 chunksize 를 지원하지 않는다.
따라서 특정 row수만큼 반복하여 끊어 읽기 위해서는 loop을 사용해야 한다.
This code shows how to whole read excel file chunk by chunk with pd.read_excel.
'''


def readList(import_file, chunk_count, chunk):
'''
이전에 읽은 rows는 건너뛰어 읽는 함수
Function for skipping old rows
'''
    print("...Skipped "+str(chunk_count*chunk)+" rows from first row")
    comp = pd.read_excel(import_path, header=0, sheet_name=0, skiprows=range(0,chunk_count*chunk), nrows=chunk)




# input 엑셀 파일 경로 (주의! 엑셀의 첫 번째 시트에 데이터가 들어있어야 한다.)
input_path = 'C:\\Users\\mypc\\Desktop\\icin\\인증개선_icin_기업명_20210106.xlsx'
# output 엑셀 파일 경로 (확장자 .xlsx는 생략할 것)
output_raw = 'C:\\Users\\mypc\\Desktop\\icin\\icin_raw_'
output_env = 'C:\\Users\\mypc\\Desktop\\icin\\icin_env_'
output_soc = 'C:\\Users\\mypc\\Desktop\\icin\\icin_soc_'
# 만료일자 기준연도
standard_year = 2020
# 끊어읽는 단위 (한번에 많이 읽어올 경우 webpage 에러가 날 수 있어 50~200이 적절함)
chunk = 50
# ---------사용자 수정 --------- #

wb = load_workbook(input_path)
sheet = wb.worksheets[0]
row_count = sheet.max_row       # 전체 row (회사명) 수
if row_count%chunk == 0:
    no_iter = row_count//chunk  # 끊어 읽기 위한 iteration 횟수
else:
    no_iter = row_count//chunk + 1
del wb
del sheet

for chunk_count in range(0,no_iter):
    print("......Reading Iteration: "+str(chunk_count+1)+"/"+str(no_iter)+"......")
    cert, comp = readList(input_path, chunk_count, chunk)
    fin_def = mergeCertInfo(cert, comp)
    exportFile(fin_def, output_raw+str(chunk_count)+".xlsx", output_env+str(chunk_count)+".xlsx", output_soc+str(chunk_count)+".xlsx", standard_year)
