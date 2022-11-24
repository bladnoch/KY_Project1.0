class Userchecker:
    def __init__(self):
        user_Checker = True
        while (user_Checker):
            print("ID :")
            user_ID = input()
            print("PW :")
            user_PW = input()

            if (user_ID == "dong5478") & (user_PW == "ehddnr0428"):
                print(user_ID, user_PW)
                user_Checker = False
            else:
                print("please insert right number")




class Menu:
    def __init__(self): #menu 생성자(Constructor)
        menu = input("1. nu customer  2. items 0. off")
        if (menu == "1") | ("nu" in menu) | ("customer" in menu):
            print("1 nu customer")
        elif (menu == "2") | ("items" in menu):
            print("2 items")
        elif (menu == "0") | ("off" in menu):
            print("off")



admin=Userchecker()
print("what you wanna do")
test=Menu()

