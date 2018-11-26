from functions import dim_aligned, i_beam

elements = []
type_elem_1 = elements.append("I-beam")
type_elem_2 = elements.append("T-beam")
type_elem_3 = elements.append("The \"Box\"")
for i, j in zip(range(1, len(elements) + 1), elements):
    print(i, ". ", j, sep='')

choice = input("Choose the element or press any key to exit: ")
if choice == str(1):
    print("You chose the I-beam element. Let's set parameters (in millimetres): ")

    l_i = int(input("Enter a length of an element: "))
    b1_i = int(input("Enter a width of a top chord: "))
    tw1_i = int(input("Enter a thickness of a top chord: "))
    b2_i = int(input("Enter a width of a bottom chord: "))
    tw2_i = int(input("Enter a thickness of a bottom chord: "))
    h_i = int(input("Enter a height of a wall: "))
    t_i = int(input("Enter a thickness of a wall: "))
    print("it's time to assign a grid of bolt holes.")

    rows = int(input("Enter a number of rows: "))
    col = int(input("Enter a number of columns: "))
    i_beam(l_i, b1_i, tw1_i, b2_i, tw2_i, h_i, t_i, rows, col)

elif choice == str(2):
    T_beam()
elif choice == str(3):
    the_box()
else:
    print("Goodbye! See you next time!")
    exit()

print("Keep working ...")
