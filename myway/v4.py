# #rules
# og_l은 모조건 왼쪽 물건의 목록을 저장하는 용도로 사용한다.
# new_l은 무조건 오른쪽 물건의 목록을 저장하는 용도로 사용한다.
#
# og_show는 왼쪽 물건의 목록을 표시하는 용도로 사용한다.
# new_show는 오른쪽 물건의 목록을 표시하는 용도로 사용한다.
#
# test.xlsx는 3종류의 목록을 3시트에 나눠서 보관한다.
#
# 왼쪽 목록을 표시할때마다 사용하는 메소드를 분리를 한다.
#
# 오른쪽 목록을 표시 할때만다 사용하는 메소드를 분리해서 만든다.
#
# 개발 주의사향
# ++메소드는 최대한 여러번 사용할 수 있어야 하고 디테일하게 나눠야 한다.
# ++화면에 보이는 상태와 뒤에서 데이터 처리를 할 때의 상태를 구분할 필요가 있음

# 구현 해야하는 개발 목록
# --4년이상 개인정보를 보호할 용도의 시트 또는 xl파일을 따로 생성.
# --왼쪽 목록의 시트를 3개로 구분해야 하고 프리셋을 넣을 가능성 높음
# --수량과 가격 입력 기능
# --날짜 추가
# --총 수납 관련 계산


# -----------
#
# og_l 규칙
# 왼쪽 시트 불러올때마다 새로운 시트로 업데이트
#     -og_l 비우기
#     -사용할 시트 결정
#     -시트 row 길이 구하기
#     -og_l에 해당 시트 저장
#     -og_show 비우기
#     -og_l 정보 og_show에 저장
#
# 오른쪽 시트 바뀔때마다 새로운 시트로 업데이트(수량 체크)
#
# new_l 규칙
# 왼쪽에서 물건 넘어올때 업데이트
# 오른쪽에서 물건 사라질 때마다 업데이트
# 불러올때 업데이트
#
# temp_l -정의 안됨
# 저장버튼이 안 눌러지면 다시 temp_l을 불러와서
# 저장이 안되면 temp_l
# 저장되면 new_l

import openpyxl
import subprocess

def setog_sheets(): #왼쪽 시트별 길이 저장 =>og_row(3개 기준)
    count = 0
    for i in range(len(og_sheets)):
        for rows in og_sheets[i].iter_rows():  # ws시트 row 길이를 count에 저장
            count += 1
        og_row[i]=count
        count=0

def setlist(): #og_l에 셀 값 저장(2개)
    row=[]
    for i in range(len(og_l)):
         for k in range(1, (og_row[i] + 1)):  # og_l[i] row 길이만큼 반복
            for j in range(1, 3):
                row.append(og_sheets[i].cell(row=k, column=j).value)
            og_l[i].append(row)
            row = []

    for i in range(len(og_l)):
        print(og_l[i])

#시트기준
#Sheet1= peersonal info devied by row
#other sheets=items by room


#빈소 특 101,102,201,202, 안치1, 안치2, 안치3
# sp101
# sp102
# sp201
# sp202
# ahn1
# ahn2
# ahn3
home = '/Users/doungukkim/Desktop/workspace/python/restinpeace/myway/excel/test.xlsx' #
room1='/Users/doungukkim/Desktop/workspace/python/restinpeace/myway/excel/room_one.xlsx'
room2='/Users/doungukkim/Desktop/workspace/python/restinpeace/myway/excel/room_two.xlsx'
room3='/Users/doungukkim/Desktop/workspace/python/restinpeace/myway/excel/room_three.xlsx'
room5='/Users/doungukkim/Desktop/workspace/python/restinpeace/myway/excel/room_five.xlsx'
room6='/Users/doungukkim/Desktop/workspace/python/restinpeace/myway/excel/room_six.xlsx'


og_file= openpyxl.load_workbook(home, data_only=True) #초기 시트 위치 저장(값으로)

og_sheets=[og_file['Sheet1'],og_file['Sheet2'],og_file['Sheet3']]  #시트 리스트에 저장 시트 이름 바꾸면 같이 바꿔야 함
og_row=['','',''] #길이 저장

og_l=[[],[],[]]
new_l=[]

og_p=[]
new_p=[]

setog_sheets()

for i in range(3):
    print(og_row[i])
setlist()