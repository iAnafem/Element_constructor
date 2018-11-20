from pyautocad import Autocad, APoint, types
import comtypes

acad = Autocad(create_if_not_exists=True)
acad.ActiveDocument.Application.Documents.Open("D:\\IT\\Bridge_truss\\pattern.dwg")
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)


def i_beam(l_i, b1_i, tw1_i, b2_i, tw2_i, h_i, t_i):
    """ This function draws the
    I-beam element in Autocad
    """
    # Creating front view
    p1 = types.aDouble(0, 0)
    p2 = types.aDouble(l_i, 0)
    p3 = types.aDouble(l_i, (tw2_i + h_i + tw1_i))
    p4 = types.aDouble(0, (tw2_i + h_i + tw1_i))
    p5 = types.aDouble(0, tw2_i)
    p6 = types.aDouble(l_i, tw2_i)
    p7 = types.aDouble(0, tw2_i + h_i)
    p8 = types.aDouble(l_i, tw2_i + h_i)
    poly_f_1 = acad.model.AddLightweightPolyline(p1 + p2 + p3 + p4 + p1)
    poly_f_2 = acad.model.AddLightweightPolyline(p5 + p6)
    poly_f_3 = acad.model.AddLightweightPolyline(p7 + p8)

    # Creating top view
    p9 = types.aDouble(0, b2_i)
    p10 = types.aDouble(l_i, b2_i)
    p11 = types.aDouble(-100, b2_i/2)
    p12 = types.aDouble(l_i+100, b2_i/2)
    poly_t_1 = acad.model.AddLightweightPolyline(p1 + p2 + p10 + p9 + p1)
    poly_t_2 = acad.model.AddLightweightPolyline(p11 + p12)

   

elements = []
type_elem_1 = elements.append("I-beam")
type_elem_2 = elements.append("T-beam")
type_elem_3 = elements.append("The \"Box\"")
for i, j in zip(range(1, len(elements) + 1), elements):
    print(i, ". ", j, sep='')

choice = input("Choose the element or press any key to exit: ")
if choice == str(1):
    print("You chose I-beam element. Let's set parameters (in millimetres): ")

    l_i = int(input("Enter length: "))
    b1_i = int(input("Enter the width of the top chord: "))
    tw1_i = int(input("Enter the thickness of the top chord: "))
    b2_i = int(input("Enter the width of the bottom chord: "))
    tw2_i = int(input("Enter the thickness of the bottom chord: "))
    h_i = int(input("Enter the height of the wall: "))
    t_i = int(input("Enter the thickness of the wall: "))

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
