import openpyxl
import array
from pyautocad import Autocad, APoint, ACAD


class Point:
    def __init__(self, coord):
        self.center = array.array('d', [coord[0], coord[1], coord[2]])
        self.x = coord[0]
        self.y = coord[1]
        self.z = coord[2]

    def offset(self, offset_x, offset_y):
        return Point([self.x + offset_x, self.y + offset_y, self.z])


class Row:
    def __init__(self, points):
        self.points = list()
        for i in points:
            self.points.append(Point(i))


class Grid:
    def __init__(self):
        self.rows = list()

    def row_len(self, row_num):
        return len(self.rows[row_num].points)

    def grid_len(self):
        result = 0
        for i in range(len(self.rows)):
            result += len(self.rows[i].points)
        return result

    def get_point(self, row, column):
        if row >= len(self.rows):
            raise ValueError('row out of range')
        if column >= len(self.rows[row].points):
            raise ValueError('column out of range')
        return self.rows[row].points[column]

    def get_upper_left(self):
        return self.get_point(0, 0)

    def get_upper_right(self):
        return self.get_point(0, len(self.rows[0].points) - 1)

    def get_down_left(self):
        return self.get_point(len(self.rows) - 1, 0)

    def get_down_right(self):
        last_row = len(self.rows) - 1
        return self.get_point(last_row, len(self.rows[last_row].points) - 1)


class Constructor:
    def __init__(self, app):
        self.app = app

    def add_line(self, point1, point2, layer):
        self.app.model.AddLine(point1, point2).layer = layer

    def add_polyline(self, group, row_num, layer):
        sequence = array.array('d', [group.get_point(row_num, 0).x,
                                     group.get_point(row_num, 0).y,
                                     group.get_point(row_num, 0).z])
        end_point = group.get_point(row_num, group.row_len(row_num) - 1).center
        for j in range(1, group.row_len(row_num)):
            sequence += group.get_point(row_num, j).center
        if len(sequence) <= 6:
            self.app.model.AddPolyline(sequence).layer = layer
        else:
            self.app.model.AddPolyline(end_point + sequence).layer = layer

    def add_block_group(self, group, block_name):
        for i in range(len(group.rows)):
            for j in range(len(group.rows[i].points)):
                self.app.model.InsertBlock(group.get_point(i, j).center, block_name, 1, 1, 1, 0)

    def add_block(self, point, block_name, scale):
        self.app.model.InsertBlock(point.center, block_name, scale, scale, scale, 0)

    def add_double_gap(self, point, scale):
        self.add_block(point.offset(0, - point.y/2), 'gap', scale)
        self.add_line(point.offset(- 1.5 * scale, - point.y / 2 - scale).center,
                      point.offset(- 1.5 * scale, -point.y - scale).center, "разрыв")
        self.add_line(point.offset(- 1.5 * scale,  - point.y/2 + scale).center,
                      point.offset(- 1.5 * scale, scale).center, "разрыв")
        self.add_line(point.offset(1.5 * scale, - point.y / 2 + scale).center,
                      point.offset(1.5 * scale, scale).center, "разрыв")
        self.add_line(point.offset(1.5 * scale, - point.y / 2 - scale).center,
                      point.offset(1.5 * scale, -point.y - scale).center, "разрыв")
        self.app.ActiveDocument.SendCommand("_trim\n\nЛ\n" + str(point.x) + "," +
                                            str(- point.y - 1) + "\n" +
                                            str(point.x) + "," +
                                            str(point.y + 1) + "\n\n\n")

    def holes_side_view(self, group1, group2):
        for i in range(len(group1.rows)):
            coord = group1.get_point(i, 0).x
            self.add_line(group2.get_point(0, 0).offset(coord, 0).center,
                          group2.get_point(0, 1).offset(coord, 0).center, 'Пунктир')
            self.add_line(group2.get_point(0, 2).offset(coord, 0).center,
                          group2.get_point(0, 3).offset(coord, 0).center, 'Пунктир')
            self.add_line(group2.get_point(2, 0).offset(coord, 0).center,
                          group2.get_point(2, 1).offset(coord, 0).center, 'Ось')


class Editor:
    def __init__(self, app):
        self.app = app

    @staticmethod
    def chain_dim(group, p1, p2):
        return str(group.row_len(p1[0]) - 1) + \
               "x" + str(int((group.get_point(p2[0], p2[1]).y -
                              group.get_point(p1[0], p1[1]).y) / (group.row_len(p1[0]) - 1))) + \
               "=" + str(group.get_point(p2[0], p2[1]).y - group.get_point(p1[0], p1[1]).y)

    def move_front_view(self, point):
        self.app.ActiveDocument.SendCommand("m\n25000,-2000\n-1000,2000\n\n" + str(point.x / 2) + ",0\n-8000,0\n")

    def move_top_view(self, point):
        self.app.ActiveDocument.SendCommand("m\n25000,-2000\n-1000,2000\n\n" + str(point.x / 2) + ",0\n-8000,-5000\n")

    def move_side_view(self):
        self.app.ActiveDocument.SendCommand("m\n10000,-2000\n-1000,2000\n\n0,0\n5000,0\n")

    def change_scale(self, scale):
        return self.app.ActiveDocument.SendCommand("CANNOSCALE\n1:" + str(scale) + "\n")

    def squeeze_left(self, point_1, point_2):
        self.app.ActiveDocument.SendCommand("_stretch\n" + str(point_1.x + 2500) + ",-2000\n" +
                                            str(point_1.x - 2500) + ",2000\n\n0,0\n" +
                                            str((point_2.x - 4500)/2) + ",0\n")

    def squeeze_right(self, point_1, point_2):
        self.app.ActiveDocument.SendCommand("_stretch\n" + str(point_1.x + 2500) + ",-2000\n" +
                                            str(point_1.x - 2500) + ",2000\n\n0,0\n" +
                                            str(-(point_2.x - 4500)/2) + ",0\n")

    def original_element(self, offset):
        self.app.ActiveDocument.SendCommand("c\n25000,-2000\n-1000,2000\n\n0,0\n" + str(offset) + "\n\n")


class Dimensions:
    def __init__(self, app):
        self.app = app

    def rotated_dim_x(self, point1, point2, start_point, indent, scale):
        loc = array.array('d', [start_point.x, start_point.y + indent * scale, start_point.z])
        dim = self.app.model.AddDimRotated(point1, point2, loc, 0)
        dim.layer = "Размер"
        return dim

    def chain_rotated_dim_x(self, group, start_point, offset, indent, scale):
        for i in range(len(group.rows) - 1):
            Dimensions(self.app).rotated_dim_x(group.get_point(i, 0).offset(0, offset).center,
                                               group.get_point(i + 1, 0).offset(0, offset).center,
                                               start_point, indent, scale)

    def rotated_dim_y(self, point1, point2, start_point, indent, scale):
        loc = array.array('d', [start_point.x + indent * scale, start_point.y, start_point.z])
        dim = self.app.model.AddDimRotated(point1, point2, loc, -1.5708)
        dim.layer = "Размер"
        return dim

    def chain_rotated_dim_y(self, group, start_point, offset, indent, scale):
        for i in range(group.row_len(0) - 1):
            Dimensions(self.app).rotated_dim_y(group.get_point(0, i).offset(offset, 0).center,
                                               group.get_point(0, i + 1).offset(offset, 0).center,
                                               start_point, indent, scale)


class Design:
    def __init__(self, app, worksheet):
        self.app = app
        self.worksheet = worksheet

    def main_header(self, scale, group):
        main_header = self.app.model.AddText(u"Конструкция "
                                             + str(self.worksheet["C2"].value) + "a "
                                             + str(self.worksheet["D2"].value)
                                             + " (1:" + str(scale) + ")",
                                             APoint(0, 0), 4 * scale)
        main_header.Alignment = ACAD.acAlignmentMiddleCenter
        main_header.TextAlignmentPoint = APoint(group.get_point(0, 1).x / 2,
                                                group.get_point(1, 2).y + 30 * scale)
        main_header.layer = "Текст заголовков"
        return main_header

    def top_header(self, text, group, scale):
        header = self.app.model.AddText(text, APoint(0, 0), 4 * scale)
        header.Alignment = ACAD.acAlignmentMiddleCenter
        header.TextAlignmentPoint = APoint(group.get_point(0, 1).x / 2,
                                           group.get_point(1, 2).y + 25 * scale)
        header.layer = "Текст заголовков"
        return header

    def side_header(self, text, group, scale):
        header = self.app.model.AddText(text, APoint(0, 0), 4 * scale)
        header.Alignment = ACAD.acAlignmentMiddleCenter
        header.TextAlignmentPoint = APoint(0, group.get_point(1, 2).y + 25 * scale)
        header.layer = "Текст заголовков"
        return header

    def insert_block(self, point, block_name, scale):
        return self.app.model.InsertBlock(point, block_name, scale, scale, scale, 0)


class Leader:
    def __init__(self, app, worksheet):
        self.app = app
        self.worksheet = worksheet

    def add_mleader_holes(self, point1, point2, group):
        text = "∅" + str(self.worksheet["M2"].value) + "\n" + str(group.grid_len()) + " отв."
        mleader = self.app.model.AddMLeader(point1 + point2, 1)
        mleader.TextString = text
        mleader.TextHeight = 2.5
        mleader.layer = "Размер"
        return mleader

    def add_mleader_pos(self, point1, point2, pos):
        mleader = self.app.model.AddMLeader(point1 + point2, 1)
        mleader.TextString = pos
        mleader.layer = "Размер"
        return mleader


class Tables:
    def __init__(self, app, worksheet):
        self.app = app
        self.worksheet = worksheet

    def add_specification(self):
        table = self.app.find_one(object_name_or_list="table")
        table.SetText(0, 0, "Спецификация " + str(self.worksheet["C2"].value) + "и "
                      + str(self.worksheet["D2"].value))
        table.SetText(18, 0, "* - масса дана с учетом 2% на сварные швы")

        for i in range(32, 40):
            for j in range(18, 22):
                table.SetCellValue(i - 30, j - 18, self.worksheet.cell(row=i, column=j).value)

        for i in range(32, 40, 2):
            table.SetCellValue(i - 30, 5, str(self.worksheet.cell(row=i, column=22).value))
            table.SetCellValue(i - 30, 6, str('%.3f' % self.worksheet.cell(row=i, column=23).value).replace('.', ','))
        table.SetCellValue(2, 7, str('%.3f' % self.worksheet.cell(row=32, column=24).value).replace('.', ','))

        for i in range(2, 18, 2):
            table.MergeCells(i, i + 1, 2, 2)
            if table.GetCellValue(i, 1) == 0:
                table.DeleteRows(i, 17 - i + 1)
                table.MergeCells(2, i - 1, 0, 0)
                table.MergeCells(2, i - 1, 7, 7)
                break

    def add_stamp(self):
        self.app.ActiveDocument.SendCommand('РЛИСТ\nТ\nГотовое\n')
        table = self.app.find_one(object_name_or_list="table")
        table.SetText(8, 6, str(self.worksheet["U58"].value))
        table.SetText(6, 8, str(self.worksheet["W55"].value))
        text = self.app.doc.PaperSpace.AddText(self.worksheet["W52"].value, APoint(0, 0), 2.5)
        text.Alignment = ACAD.acAlignmentMiddleCenter
        text.TextAlignmentPoint = APoint(520, 229.50)


def get_coordinates(point_group, range_, worksheet):
    for i in range(range_[0], range_[1] + 1):
        coords = []
        for j in range(range_[2], range_[3] + 1, 3):
            coord = [
                worksheet.cell(row=i, column=j).value,
                worksheet.cell(row=i, column=j + 1).value,
                worksheet.cell(row=i, column=j + 2).value,
            ]
            if "None" in coord:
                continue
            coords.append(coord)
        if len(coords) != 0:
            point_group.rows.append(Row(coords))
    return point_group


def get_autocad(path):
    app = Autocad(create_if_not_exists=True)
    app.Application.Documents.Open(path)
    app.prompt('Hello, Autocad from Python\n')
    print(app.doc.Name)
    return app


def get_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb


def autocad_save(app, path, worksheet):
    app.ActiveDocument.SaveAs(path +
                              "_" +
                              str(worksheet["W52"].value) +
                              "_" +
                              str(worksheet["U58"].value)
                              )
