import os
from pathlib import Path
import openpyxl
import os.path
from openpyxl.worksheet.table import Table, TableStyleInfo
import tkinter.ttk
import tkinter as tk
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QBoxLayout
from PyQt5.QtCore import Qt
import sys
from PyQt5.QtWidgets import * # PyQt5, Designer, tool 인스톨하셈
from PyQt5 import uic
#from PyQt5.QAxContaniner import *
from PyQt5.QtGui import *
import subprocess # 외부프로그램 여는 라이브러리
import pandas as pd # xlrd, pandas 인스톨하셈


from_class = uic.loadUiType("garam.ui")[0]
#QMainWindo, from_class를 상속받아 내 클래스 MyWindow
class MyWindow(QMainWindow, from_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self) #this가 상속받은곳에 없으면 부모를 찾아감

    def setUI(self):
        self.setupUi(self) # UI 관련 함수
        self.print_bt.clicked.connect(self.print)

    def table1(self):
        self.tb1.setItem(0, 0, QTableWidgetItem('고무장갑'))
        self.tb1.setItem(0, 1, QTableWidgetItem('1000'))
        self.tb1.setItem(0, 2, QTableWidgetItem('93'))
        self.tb1.setItem(1, 0, QTableWidgetItem('라이터'))
        self.tb1.setItem(1, 1, QTableWidgetItem('500'))
        self.tb1.setItem(1, 2, QTableWidgetItem('13'))
        self.tb1.setItem(2, 0, QTableWidgetItem('지우개'))
        self.tb1.setItem(2, 1, QTableWidgetItem('300'))
        self.tb1.setItem(2, 2, QTableWidgetItem('122'))

    def table2(self):
        self.tb1.setItem(0, 0, QTableWidgetItem('고무장갑'))
        self.tb1.setItem(0, 1, QTableWidgetItem('1000'))
        self.tb1.setItem(0, 2, QTableWidgetItem('93'))
        self.tb1.setItem(1, 0, QTableWidgetItem('라이터'))
        self.tb1.setItem(1, 1, QTableWidgetItem('500'))
        self.tb1.setItem(1, 2, QTableWidgetItem('13'))
        self.tb1.setItem(2, 0, QTableWidgetItem('지우개'))
        self.tb1.setItem(2, 1, QTableWidgetItem('300'))
        self.tb1.setItem(2, 2, QTableWidgetItem('122'))

    def load(self):
        pass

    # 외부 프로그램 실행
    def print(self):
        subprocess.run('C:/Garam/Print/WinFormsApp11.exe')

    # 재고관리 엑셀 실행
    def item(self):
        pass
        #subprocess.run('C:/Garam/Print/List.xlsx')

    def list1(self):
        pd.read_excel('List.xlsx')


if __name__ == "__main__":

    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()