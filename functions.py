from pyautocad import Autocad, APoint, types
import comtypes
import array
import openpyxl

wb = openpyxl.load_workbook("D:\\IT\\Bridge_truss\\settings.xlsx", data_only=True)
sheet = wb["Sheet"]

acad = Autocad(create_if_not_exists=True)
# acad.Application.Documents.Open("D:\\IT\\Bridge_truss\\pattern.dwg")
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)


def dim_aligned(point1, point2, scale, indent_x=0, indent_y=0, start_point_x=0, start_point_y=0, move_x=0, move_y=0):
    """ This function places an aligned dimension.
        Parameters:
        1. point1, point2 - the points between which you want to place a dimension.
        2. scale - a scale you want to create the current dimension.
        3. indent_x, indent_y - indents in accordance with GOST:
          - 10 mm for first step;
          - 8 mm for other steps.
          For 2-nd line of dimensions indent = 10+8, for 3-rd -10+8+8, etc.

          !!! BE CAREFUL !!! with "+" and "-" !!!

        4. start_point_x, start_point_y  - x and y-coordinates from which you need to make an indent.
        5. move_x, move_y - view's offset.
    """
    # It's measured from 0, 0, 0 !!!
    loc = array.array('d', [start_point_x + indent_x * scale, start_point_y + indent_y * scale, 0])
    # ==============================
    dim = acad.model.AddDimAligned(point1, point2, loc)
    dim.layer = "Размер"
    if move_x != 0 or move_y != 0:
        dim.move(APoint(0, 0), APoint(move_x, move_y))


def holes(diameter: str, start_point_x, start_point_y, holes_grid, move_x=0, move_y=0):

    for x in range(0, len(holes_grid), 2):
        if holes_grid[x] != 0:
            block = acad.model.InsertBlock(APoint(start_point_x + holes_grid[x],
                                                  start_point_y + holes_grid[x+1]),
                                           "hole_" + diameter, 1, 1, 1, 0)
            block.move(APoint(0, 0), APoint(move_x, move_y))


def i_beam(l_i, b1_i, tw1_i, b2_i, tw2_i, h_i, t_i, rows, col):
    """ This function creates the geometry
        of the I-beam element in Autocad
    """

    # Creating front view.

    scale_front = int(input("Enter a scale of the front view: "))
    print("Creating the front view...")
    p1 = array.array('d', [0, 0, 0])
    p2 = array.array('d', [l_i, 0, 0])
    p3 = array.array('d', [l_i, tw2_i, 0])
    p4 = array.array('d', [0, tw2_i, 0])
    p5 = array.array('d', [0, tw2_i + h_i, 0])
    p6 = array.array('d', [l_i, tw2_i + h_i, 0])
    p7 = array.array('d', [l_i, (tw2_i + h_i + tw1_i), 0])
    p8 = array.array('d', [0, (tw2_i + h_i + tw1_i), 0])
    poly_f_b1 = acad.model.AddPolyline(p1 + p2 + p3 + p4 + p1)
    poly_f_b2 = acad.model.AddPolyline(p5 + p6 + p7 + p8 + p5)
    poly_f_h1 = acad.model.AddPolyline(p3 + p4 + p5 + p6 + p3)
    poly_f_b1.layer = "Основная 0.25"
    poly_f_b2.layer = "Основная 0.25"
    poly_f_h1.layer = "Основная 0.25"
    dim_aligned(p1, p2, scale_front, indent_y=-10)
    dim_aligned(p2, p3, scale_front, indent_x=10, start_point_x=l_i)

    # Creating top view

    scale_top = int(input("Enter a scale of the top view: "))
    if scale_front != scale_top:
        input("Please change a scale in model and press \"Y\": ")
    offset_view_y = -5000
    print("Creating the top view...")
    p9 = array.array('d', [0, b2_i, 0])
    p10 = array.array('d', [l_i, b2_i, 0])
    p11 = array.array('d', [-100, b2_i / 2, 0])
    p12 = array.array('d', [l_i + 100, b2_i / 2, 0])

    poly_t_1 = acad.model.AddPolyline(p1 + p2 + p10 + p9 + p1)
    poly_t_2 = acad.model.AddPolyline(p11 + p12)

    poly_t_1.layer = "Основная 0.25"
    poly_t_2.layer = "Ось"

    poly_t_1.move(APoint(0, 0), APoint(0, offset_view_y))
    poly_t_2.move(APoint(0, 0), APoint(0, offset_view_y))

    holes_grid_1 = []
    for i in range(7, 7+col):
        for j in range(2, rows*2+2):
            holes_grid_1.append(sheet.cell(row=i, column=j).value)
    # holes_grid2 = ((50,	130), (130,	130), (210,	370), (290,	290))
    # holes_grid3 = ((50, 290), (130, 290))
    # holes_grid4 = ((50, 370), (130, 370))
    holes("25", 0, 0, holes_grid_1, move_y=offset_view_y)
    # holes("25", 0, 0, holes_grid2, move_y=offset_view_y)
    # holes("25", 0, 0, holes_grid3, move_y=offset_view_y)
    # holes("25", 0, 0, holes_grid4, move_y=offset_view_y)
    dim_aligned(p1, p2, scale_top, indent_y=-10, move_y=offset_view_y)
# acad.ActiveDocument.SaveAs('D:\\IT\\Bridge_truss\\1.dwg')
