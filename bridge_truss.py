import comtypes
import array
from pyautocad import Autocad, APoint, types
from functions import dim_aligned, i_beam

acad = Autocad(create_if_not_exists=True)
# acad.ActiveDocument.Application.Documents.Open("D:\\IT\\Bridge_truss\\pattern.dwg")
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)

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
    i_beam(l_i, b1_i, tw1_i, b2_i, tw2_i, h_i, t_i)
elif choice == str(2):
    T_beam()
elif choice == str(3):
    the_box()
else:
    print("Goodbye! See you next time!")
    exit()
# acad = Autocad(create_if_not_exists=True)
# acad.ActiveDocument.Application.Documents.Open("D:\\IT\\Bridge_truss\\pattern.dwg")
# acad.prompt("Hello, Autocad from Python\n")
# print(acad.doc.Name)


# p1 = APoint(0, 0)
# p2 = APoint(50, 25)

# acad.model.AddLine(p1, p2)
# acad.ActiveDocument.SaveAs('D:\\IT\\Bridge_truss\\1.dwg')

print("Keep working ...")
