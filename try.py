from pyautocad import Autocad, APoint, types
import comtypes
import array
acad = Autocad(create_if_not_exists=True)
# acad.ActiveDocument.Application.Documents.Open("D:\\IT\\Bridge_truss\\pattern.dwg")
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)


p1 = array.array('d', [0, 0, 0])
p2 = array.array('d', [100, 0, 0])
p3 = array.array('d', [100, 350, 0])
p4 = array.array('d', [0, 350, 0])

poly_f_1 = acad.model.AddPolyline(p1 + p2 + p3 + p4 + p1)

block = acad.model.InsertBlock(APoint(0, 500), "round", 10, 10, 10, 0)
block.layer = "layer1"

acad.model.AddLeader()

loc_x = array.array('d', [0, -100, 0])
loc_y = array.array('d', [200, 0, 0])

dim_1 = acad.model.AddDimAligned(p1, p2, loc_x).move(APoint(0, 100), APoint(0, 500))
dim_2 = acad.model.AddDimAligned(p2, p3, loc_y).move(APoint(0, 100), APoint(0, 500))

poly_f_1.move(APoint(0, 100), APoint(0, 500))

poly_f_1.layer = "layer1"
