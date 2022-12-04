from tkinter import * # tkinter의 모든 함수 가져오기
from tkinter import messagebox, filedialog
import subprocess
import openpyxl
import os
from openpyxl.worksheet.table import Table, TableStyleInfo
import tkinter.ttk


# def openfile():
def openxl():
    def close():
        openxl.quit()
        openxl.destroy()

    openxl=Tk()

    openxl.geometry("300x170")  # 창의 크기
    openxl.title("장례식장 재고관리 프로그램 Ver1.221123")  # 창의 제목
    openxl.option_add("*Font", "맑은고딕 11")  # 전체 폰트

    l_item = Label(openxl)
    l_item.config(text="물품명", width=10, relief="solid")
    l_item.place(x=20,y=20)

    l_price = Label(openxl)
    l_price.config(text="단가", width=10, relief="solid")
    l_price.place(x=20, y=50)

    l_count = Label(openxl)
    l_count.config(text="단위", width=10, relief="solid")
    l_count.place(x=20, y=80)

    e_item = Entry(openxl)
    e_item.config(width=20, relief="solid", borderwidth=2)

    e_price = Entry(openxl)
    e_price.config(width=20, relief="solid", borderwidth=2)

    e_count = Entry(openxl)
    e_count.config(width=20, relief="solid", borderwidth=2)

    e_item.place(x=110,y=20)
    e_price.place(x=110,y=50)
    e_count.place(x=110,y=80)

    save = Button(openxl, text="저장")
    save.config(width=6, height=2)
    save.place(x=60,y=115)

    cancel = Button(openxl, text="취소")
    cancel.config(width=6, height=2)
    cancel.place(x=150,y=115)

    openxl.mainloop()

def close():
    win.quit()
    win.destroy()

def first():
    row=[]
    count=0

    for rows in ws.iter_rows(): #기본 물품의 rows 값
        count += 1

    for i in range(1, (count + 1)):  # og_list에 기본 물품 저장
        for j in range(1, 8):
            row.append(ws.cell(row=i, column=j).value)
        og_l.append(row)
        row = []

    for i in range(1,count):
        리스트.insert(i-1,og_l[i][2])

def listitem():

    messagebox(리스트.curselection())



##################################################   global variable   ##########################

home = '/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/test.xlsx' #기본 물품 엑셀 위치 저장

room1='/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/room_one.xlsx'
room2='/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/room_two.xlsx'
room3='/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/room_three.xlsx'
room4='/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/room_four.xlsx'
room5='/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/room_five.xlsx'
room6='/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/room_six.xlsx'

wb= openpyxl.load_workbook(home, data_only=True) #초기 시트 위치 저장(값으로)
ws=wb['Sheet1'] #초기 시트 사용 선언

global og_l #초기 리스트 저장공간
og_l=[]
global new_l #새로운 리스트 저장공간

global tree #make tree
global treelist #list
treelist=[]




##################################################   tkinter   ##########################
win = Tk() # 창 생성
win.geometry("1000x720") # 창의 크기
win.title("장례식장 재고관리 프로그램 Ver1.221123") # 창의 제목
win.option_add("*Font", "맑은고딕 11") # 전체 폰트
#win.resizable(False, False) #윈도우 사이즈 조절 불가

ID_lab = Label(win)
ID_lab.config(text = "ID", width=10, relief="solid")
물품명 = Label(win)
물품명.config(text = "물품명", width=15, relief="solid",borderwidth=0)
고인명_lab = Label(win)
고인명_lab.config(text = "고인명",width=10, relief="solid")
상주명_lab = Label(win)
상주명_lab.config(text = "상주명",width=10, relief="solid")
빈소_lab = Label(win)
빈소_lab.config(text = "빈소", width=10, relief="solid")

수납금액_lab = Label(win)
수납금액_lab.config(text = "수납금액", width=10, relief="solid")
받은금액_lab = Label(win)
받은금액_lab.config(text = "받은금액", width=10, relief="solid")
거스름돈_lab = Label(win)
거스름돈_lab.config(text = "거스름돈", width=10, relief="solid")


##################################################   entry   ##########################
ID = Entry(win)
ID.config(width=10,relief="solid",borderwidth=2)
고인명 = Entry(win)
고인명.config(width=10,relief="solid",borderwidth=2)
상주명 = Entry(win)
상주명.config(width=10,relief="solid",borderwidth=2)
빈소 = Entry(win)
빈소.config(width=30,relief="solid",borderwidth=2)

수납금액 = Entry(win)
수납금액.config(width=20,relief="solid",borderwidth=2)
받은금액 = Entry(win)
받은금액.config(width=20,relief="solid",borderwidth=2)
거스름돈 = Entry(win)
거스름돈.config(width=20,relief="solid",borderwidth=2)

리스트 = Listbox(win, selectmode = 'extended',width = 15, height = 30)
리스트.yview()

##################################################   buttons   ##########################
저장 = Button(win, text = "저장")
저장.config(width=10,height=2)
불러오기 = Button(win, text = "불러오기")
불러오기.config(width=10,height=2)
닫기 = Button(win, text = "닫기",command=close)
닫기.config(width=10,height=3)
물품수정 = Button(win, text = "풀품수정",command=openxl)
물품수정.config(width=7,height=2)


삭제 = Button(win, text = "삭제")
삭제.config(width=10,height=3)
Set = Button(win, text = "checker")
Set.config(width=10,height=3)

first()

##################################################   place   ##########################
#labels
ID_lab.place(x=10,y=10)
고인명_lab.place(x=210,y=10)
상주명_lab.place(x=410,y=10)
빈소_lab.place(x=10,y=50)
수납금액_lab.place(x=620,y=10)
받은금액_lab.place(x=620,y=60)
거스름돈_lab.place(x=620,y=110)

#엔트리 위치
ID.place(x=110,y=10)
고인명.place(x=310,y=10)
상주명.place(x=510,y=10)
빈소.place(x=110,y=50)


수납금액.place(x=720,y=10)
받은금액.place(x=720,y=60)
거스름돈.place(x=720,y=110)

#버튼 위치
물품명.place(x=-15,y=220)
저장.place(x= 350, y=50)
불러오기.place(x=470,y=50)
닫기.place(x=370, y=150)
물품수정.place(x=10, y=150)
삭제.place(x=130, y=150)
Set.place(x=250, y=150)

리스트.place(x=20, y=240)

win.mainloop() # 창 실행