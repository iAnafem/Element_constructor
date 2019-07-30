import array
from pyautocad import APoint, ACAD


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