from tkinter import * # tkinter의 모든 함수 가져오기
import tkinter as tk
import tkinter.ttk
import tkinter.ttk as ttk
import openpyxl
import pandas as pd
#from datetime improt datetime

'''
pip install pandas
pip install xlrd
pip install openpyxl
'''

#함수 정의 부분
#def time():

#def ID_a():

def close():
    win.quit()
    win.destroy()

#Tkinter 윈도우 화면
win = Tk() # 창 생성
win.geometry("1600x900") # 창의 크기
win.title("장례식장 재고관리 프로그램 Ver1.221123") # 창의 제목
win.option_add("*Font", "맑은고딕 13") # 전체 폰트
#win.resizable(False, False) #윈도우 사이즈 조절 불가
tab_bt=tkinter.ttk.Notebook(win, width=300, height=630)

#########################   excel   ##########################

'''
#########################    menu   ##########################

menu = Menu(win)

menu_1 = Menu(menu, tearoff = 0)
menu_1.add_command(label = "로그인")
menu_1.add_command(label = "인쇄")
menu_1.add_command(label = "장부")
menu_1.add_separator()
menu_1.add_command(label = "종료", command = close)
menu.add_cascade(label = "메뉴", menu = menu_1)

win.config(menu = menu)
'''

#########################   config  ##########################



#레이블 정의
ID_lab = Label(win)
ID_lab.config(text = "ID",width=10, relief="solid")
고인명_lab = Label(win)
고인명_lab.config(text = "고인명",width=10, relief="solid")
상주명_lab = Label(win)
상주명_lab.config(text = "상주명",width=10, relief="solid")
빈소_lab = Label(win)
빈소_lab.config(text = "빈소", width=10, relief="solid")
빈소기간_lab = Label(win)
빈소기간_lab.config(text = "빈소기간", width=10, relief="solid")
안치기간_lab = Label(win)
안치기간_lab.config(text = "안치기간", width=10, relief="solid")
물결1_lab = Label(win)
물결1_lab.config(text = "~", width=10)
물결2_lab = Label(win)
물결2_lab.config(text = "~", width=10)
###
수납금액_lab = Label(win)
수납금액_lab.config(text = "수납금액", width=10, relief="solid")
받은금액_lab = Label(win)
받은금액_lab.config(text = "받은금액", width=10, relief="solid")
거스름돈_lab = Label(win)
거스름돈_lab.config(text = "거스름돈", width=10, relief="solid")

#엔트리 정의
ID = Entry(win)
ID.config(width=10,relief="solid",borderwidth=2)
고인명 = Entry(win)
고인명.config(width=10,relief="solid",borderwidth=2)
상주명 = Entry(win)
상주명.config(width=10,relief="solid",borderwidth=2)
빈소 = Entry(win)
빈소.config(width=60,relief="solid",borderwidth=2)
빈소기간1 = Entry(win)
빈소기간1.config(width=20,relief="solid",borderwidth=2)
안치기간1 = Entry(win)
안치기간1.config(width=20,relief="solid",borderwidth=2)
빈소기간2 = Entry(win)
빈소기간2.config(width=20,relief="solid",borderwidth=2)
안치기간2 = Entry(win)
안치기간2.config(width=20,relief="solid",borderwidth=2)
수납금액 = Entry(win)
수납금액.config(width=20,relief="solid",borderwidth=2)
받은금액 = Entry(win)
받은금액.config(width=20,relief="solid",borderwidth=2)
거스름돈 = Entry(win)
거스름돈.config(width=20,relief="solid",borderwidth=2)

#버튼 정의
저장 = Button(win, text = "저장")
저장.config(width=10,height=2)
불러오기 = Button(win, text = "불러오기")
불러오기.config(width=10,height=2)
#btn.config(command=ID_a)
결제 = Button(win, text = "결제")
결제.config(width=10,height=2)
닫기 = Button(win, text = "닫기")
닫기.config(width=10,height=2,command =close)
재고관리 = Button(win, text = "재고관리")
재고관리.config(width=10,height=2)

'''
식당판매 = Button(win, text = "식당판매")
식당판매.config(width=10,height=3)
매점판매 = Button(win, text = "매점판매")
매점판매.config(width=10,height=3)
Set = Button(win, text = "기본 Set")
Set.config(width=10,height=3)
'''

#탭 정의
tab1=tkinter.Frame(win)
tab_bt.add(tab1, text="식당판매")
tree1 = ttk.Treeview(tab1, columns=(1, 2, 3), height=500, show="headings")
tree1.place(x=0,y=0)
# 필드명
tree1.heading(1, text="물품")
tree1.heading(2, text="단가")
tree1.heading(3, text="갯수")
# 기본 너비
tree1.column(1, width=93)
tree1.column(2, width=93)
tree1.column(3, width=93)
# 테이블 스크롤바 표시
scroll = ttk.Scrollbar(tab1, orient="vertical", command=tree1.yview)
scroll.pack(side='right', fill='y')
tree1.configure(yscrollcommand=scroll.set)
# 기본 데이터 추가
df = pd.read_excel("List.xlsx", engine = "openpyxl", sheet_name="매점물품")


#for val in data:
    #tree1.insert('', 'end', values=(val[0], val[1], val[2]))

tab2=tkinter.Frame(win)
tab_bt.add(tab2, text="매점판매")
tree2 = ttk.Treeview(tab2, columns=(1, 2, 3), height=30, show="headings")
tree2.place(x=0,y=0)
# 필드명
tree2.heading(1, text="물품")
tree2.heading(2, text="단가")
tree2.heading(3, text="갯수")
# 기본 너비
tree2.column(1, width=95)
tree2.column(2, width=95)
tree2.column(3, width=95)
# 테이블 스크롤바 표시
scroll = ttk.Scrollbar(tab2, orient="vertical", command=tree2.yview)
scroll.pack(side='right', fill='y')
tree2.configure(yscrollcommand=scroll.set)
# 기본 데이터 추가
data = [["11", "12", "13"],
        ["4", "5", "6"],
        ["7", "8", "9"],
        ["10", "11", "12"],
        ["13", "14", "15"],
        ["16", "17", "18"]]

for val in data:
    tree2.insert('', 'end', values=(val[0], val[1], val[2]))

tab3=tkinter.Frame(win)
tab_bt.add(tab3, text="장의용품")

tab4=tkinter.Frame(win)
tab_bt.add(tab4, text="임시버튼")

tab5=tkinter.Frame(win)
tab_bt.add(tab5, text="임시버튼")
'''
리스트 = Listbox(win, selectmode = 'extended',width = 40, height = 36,)
리스트.yview()
리스트.insert(0, "1번")
리스트.insert(1, "2번") #반복문으로 딕셔너리, 튜플, 리스트 사용 가
'''

#########################   place  ##########################

#레이블 위치
ID_lab.place(x=10,y=10)
고인명_lab.place(x=210,y=10)
상주명_lab.place(x=410,y=10)
빈소_lab.place(x=10,y=50)
빈소기간_lab.place(x=10,y=90)
안치기간_lab.place(x=10,y=130)
물결1_lab.place(x=250,y=90)
물결2_lab.place(x=250,y=130)

수납금액_lab.place(x=620,y=10)
받은금액_lab.place(x=620,y=60)
거스름돈_lab.place(x=620,y=110)

#엔트리 위치
ID.place(x=110,y=10)
고인명.place(x=310,y=10)
상주명.place(x=510,y=10)
빈소.place(x=110,y=50)
빈소기간1.place(x=110,y=90)
안치기간1.place(x=110,y=130)
빈소기간2.place(x=310,y=90)
안치기간2.place(x=310,y=130)

수납금액.place(x=720,y=10)
받은금액.place(x=720,y=60)
거스름돈.place(x=720,y=110)

#버튼 위치
저장.place(x= 500, y=80)
불러오기.place(x= 500, y=130)
결제.place(x=900, y=10)
닫기.place(x=900, y=80)
재고관리.place(x=10, y=170)
#식당판매.place(x=10, y=170)
#매점판매.place(x=120, y=170)
#Set.place(x=240, y=170)

#탭위치
tab_bt.place(x=10,y=230)

#리스트 위치
#리스트.place(x=10, y=230)

win.mainloop() # 창 실행