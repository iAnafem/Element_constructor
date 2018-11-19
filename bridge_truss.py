
from pyautocad import Autocad, APoint
import comtypes

acad = Autocad(create_if_not_exists=True, visible=True)
dwgName = "D:\\IT\\Bridge_truss\\pattern.dwg"
acad.ActiveDocument.Application.Documents.Open(dwgName)
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)


p1 = APoint(0, 0)
p2 = APoint(50, 25)

acad.model.AddLine(p1, p2)

#acad.ActiveDocument.SaveAs('D:\\IT\\Bridge_truss\\1.dwg')