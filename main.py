import os

# class Admin:
#     def __init__(self):
#         user_Check = True
#         while (user_Check):
#             print("ID :")
#             user_ID = input()
#             print("PW :")
#             user_PW = input()
#
#             if (user_ID == "dong5478") & (user_PW == "ehddnr0428"):
#                 print(user_ID, user_PW)
#                 user_Check = False
#             else:
#                 print("wrong info")
#                 print("")

# class Menu:
#     def __init__(self): #menu 생성자(Constructor)
#         print("what you wanna do")
#         menu = input("1. nu customer  2. items 0. off")
#         if (menu == "1") | ("nu" in menu) | ("customer" in menu):
#             print("1 nu customer")
#         elif (menu == "2") | ("items" in menu):
#             print("2 items")
#         elif (menu == "0") | ("off" in menu):
#             print("off")

class Money:
    pay=0
    payed=0
    change=0
    def __init__(self):
        Money.pay=1000
        Money.payed=100
        Money.change=900
        print("수납금액 :",Money.pay)
        print("받은금액 :", Money.payed)
        print("거스름돈 :", Money.change)

if __name__ == '__main__':
    # admin=Admin()
    # menu=Menu()
    money=Money()

