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


admin=Userchecker()

# user_Checker=True
# while(user_Checker):
#     print("ID :")
#     user_ID = input()
#     print("PW :")
#     user_PW = input()
#
#     if (user_ID=="dong5478") & (user_PW=="ehddnr0428"):
#         print(user_ID, user_PW)
#         user_Checker=False
#     else:
#         print("please insert right number")