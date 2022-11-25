# id : dong
# pw : 1234

# 로그인
class Admin:
    def __init__(self):
        user_Check = True
        while (user_Check):
            print("ID :")
            user_ID = input()
            print("PW :")
            user_PW = input()



            if (user_ID == "dong") & (user_PW == "1234"):
                print(user_ID, user_PW)
                user_Check = False
            else:
                print("wrong info")
                print("")
#계산
class Money:
    pay=1000
    payed=100
    change=900
    def setPay(self, ppay):
        Money.pay=ppay
    def setPayed(self,ppayed):
        Money.payed=ppayed
    def setChange(self,cchange):
        Money.change=cchange



#메뉴
class Menu:
    tf=True
    def __init__(self): #menu 생성자(Constructor)

        while(Menu.tf):
            print(">>----------------------------")
            print("id","고인명","상주명","빈소")
            print("빈소기간","안치기간")
            print("수닙금액 :",Money.pay,"   받은 :",Money.payed,"   거스름돈 :",Money.change)
            print("items")
            print("------------------------------")

            menu = input("1. nu customer  2. items 3. how much 4. over \n")
            print("------------------------------")
            if (menu == "1") | ("nu" in menu) | ("customer" in menu):
                print("1 nu customer")
            elif (menu == "2") | ("items" in menu):
                print("2 items")
            elif (menu == "3") | ("how" in menu)| ("much" in menu):
                print("수납 금액 : ",Money.pay)
                self.payed=int(input("받은금액 : "))
                self.Cm=Money()
                self.Cm.setPayed(self.payed)
                print("거스름돈 : ", Money.pay-Money.payed)
                self.Cm.setChange(Money.pay-Money.payed)
            else:
                print("log off...")
                Menu.tf=False
        print("")


if __name__ == '__main__':
    # admin=Admin()
    menu=Menu()
    money=Money()

