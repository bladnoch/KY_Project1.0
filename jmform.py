from tkinter import * # tkinter의 모든 함수 가져오기
import openpyxl

#from datetime improt datetime

#함수 정의 부분
#def time():

#def ID_a():



# class insertData:
#     def __init__(self):  # menu 생성자(Constructor)


class form:
    def start():
        def close():
            win.quit()
            win.destroy()

        #Tkinter 윈도우 화면
        win = Tk() # 창 생성
        win.geometry("1000x720") # 창의 크기
        win.title("장례식장 재고관리 프로그램 Ver1.221123") # 창의 제목
        win.option_add("*Font", "맑은고딕 11") # 전체 폰트


        llist = Listbox(win, selectmode = 'extended',width = 122, height = 30)
        llist.insert(0, "1번")
        llist.insert(1,"jjkj")

        #리스트 위치
        llist.place(x=10, y=210)


        win.mainloop() # 창 실행

    if __name__ == '__main__':

        start();
        # call.setList("hello")
        # 주소 변수에 저장
        home = "/Users/doungukkim/Desktop/workspace/python/restinpeace/excelhere/test.xlsx"

        # 엑셀 불러오기
        wb = openpyxl.load_workbook(home)

        # 엑셀 시트 선택
        ws = wb['Sheet1']