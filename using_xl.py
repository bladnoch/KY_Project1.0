import openpyxl

#주소 변수에 저장
home="/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/test.xlsx"
# 엑셀 불러오기
wb= openpyxl.load_workbook(home)

# 엑셀 시트 선택
ws=wb['Sheet1']

# print(ws['A1'].value)
#데이터 수정하기

#row 기준으로 출력
# for rows in ws.iter_rows():
#     for cell in rows:
#
#         print(cell.value, end="\t\t\t")
#     print("")

##F1~F40까지를 0으로 수정
test="F"
ws['F1']="수량"
for i in range (2,41):
    test2=test+str(i)
    ws[test2]=0
#ws['A3']= "X"
# ws['A1']= "XXXXXXX"
# ws['C3']="테스트 숴정"
# ws['D3']= "X"
# ws['E3']= "XXX"
# ws['F3']= ""
# ws['G3']= "X"

#엑셀 저장
# wb.save(home)