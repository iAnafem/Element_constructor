from functions import i_beam, choose_element
import openpyxl

print("""Welcome to the \"Bridge_truss\" module !!!`
You can draw some blueprints using this module.
You can get initial data from MS Excel or type 
necessary parameters right now.
""")

import_type = input("Would you like to use excel Y/N ?")
if import_type.lower() == "y":
    wb = openpyxl.load_workbook("D:\\IT\\Bridge_truss\\settings.xlsx", data_only=True)
    sheet_name = input("Please enter the name of the sheet which you want to open to receive an initial data: ")
    sheet = wb[sheet_name]
    choose_element()
    choice = input("First, choose the element or press any key to exit: ")
    if choice == str(1):
        print("You chose the I-beam element. Processing initial data...")


        i_beam(l_i, b1_i, tw1_i, b2_i, tw2_i, h_i, t_i, sheet_name)


# print("Well, than let's enter some parameters by hands ")

# if choice == str(1):
#    print("You chose the I-beam element. Let's set parameters (in millimetres): ")

#    l_i = int(input("Enter a length of an element: "))
#    b1_i = int(input("Enter a width of a top chord: "))
#    tw1_i = int(input("Enter a thickness of a top chord: "))
#    b2_i = int(input("Enter a width of a bottom chord: "))
#    tw2_i = int(input("Enter a thickness of a bottom chord: "))
#    h_i = int(input("Enter a height of a wall: "))
#    t_i = int(input("Enter a thickness of a wall: "))

#    i_beam(l_i, b1_i, tw1_i, b2_i, tw2_i, h_i, t_i, sheet_name)

elif choice == str(2):
    T_beam()
elif choice == str(3):
    the_box()
else:
    print("Goodbye! See you next time!")
    exit()

print("Keep working ...")
