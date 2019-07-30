import array


class Constructor:
    """
    This is a class that represents the main AutoCAD drawing actions
    """
    def __init__(self, app):
        """
        The constructor for Constructor class.
        :param app: an instance of AutoCAD application.
        """
        self.app = app

    def add_line(self, point1, point2, layer):
        """
        This method draws AutoCAD object Line between two specified points and in the specified layer.
        :param point1: the starting point of the line;
        :param point2: the endpoint of the line;
        :param layer: desired layer;
        """
        self.app.model.AddLine(point1, point2).layer = layer

    def add_polyline(self, group, row_num, layer):
        """
        This method draws AutoCAD object Polyline between any number of specified points and in the specified layer.
        :param group: Grid class object;
        :param row_num: number of the row;
        :param layer: desired layer;
        """
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
        """
        This method inserts the specified AutoCAD Block object to every point in the specified group of points.
        :param group: Grid class object;
        :param block_name: name of the inserted block.
        """
        for i in range(len(group.rows)):
            for j in range(len(group.rows[i].points)):
                self.app.model.InsertBlock(group.get_point(i, j).center, block_name, 1, 1, 1, 0)

    def add_block(self, point, block_name, scale):
        """
        This method inserts the specified AutoCAD Block object to the specified point.
        :param point: Point class object;
        :param block_name: name of the inserted block.
        :param scale: desired scale.
        """
        self.app.model.InsertBlock(point.center, block_name, scale, scale, scale, 0)

    def add_double_gap(self, point, scale):
        """
        This method inserts the double gap lines to the specified point.
        :param point: Point class object;
        :param scale: desired scale
        """
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
        """
        This method creates the side view of holes of the specified groups.
        :param group1: Grid class object;
        :param group2: Grid class object;
        """
        for i in range(len(group1.rows)):
            coord = group1.get_point(i, 0).x
            self.add_line(group2.get_point(0, 0).offset(coord, 0).center,
                          group2.get_point(0, 1).offset(coord, 0).center, 'Пунктир')
            self.add_line(group2.get_point(0, 2).offset(coord, 0).center,
                          group2.get_point(0, 3).offset(coord, 0).center, 'Пунктир')
            self.add_line(group2.get_point(2, 0).offset(coord, 0).center,
                          group2.get_point(2, 1).offset(coord, 0).center, 'Ось')


class Editor:
    """
    This is a class that represents the main AutoCAD editing actions.
    """
    def __init__(self, app):
        """
        The constructor for Editor class.
        :param app: an instance of AutoCAD application.
        """
        self.app = app

    @staticmethod
    def chain_dim(group, p1, p2) -> str:
        """
        This method edits the text on the dimension leader. Use it, if you need to specify the chain of
        dimensions instead of the one real dimension, for example, between the starting point and the endpoint
        in the row of holes.
        :param group: Grid class object;
        :param p1: the starting point;
        :param p2: the endpoint;
        :return:
        """
        return str(group.row_len(p1[0]) - 1) + "x" + str(int(
            (group.get_point(p2[0], p2[1]).y - group.get_point(p1[0], p1[1]).y) / (
                    group.row_len(p1[0]) - 1))) + "=" + str(
            group.get_point(p2[0], p2[1]).y - group.get_point(p1[0], p1[1]).y)

    def move_front_view(self, point):
        """
        This method moves the front view for properly location inside the blueprint, via AutoCAD console;
        :param point: Point class object, which we need to 'hook' the view;
        """
        self.app.ActiveDocument.SendCommand("m\n25000,-2000\n-1000,2000\n\n" + str(point.x / 2) + ",0\n-8000,0\n")

    def move_top_view(self, point):
        """
        This method moves the top view for properly location inside the blueprint, via AutoCAD console;
        :param point: Point class object, which we need to 'hook' the view;
        """
        self.app.ActiveDocument.SendCommand("m\n25000,-2000\n-1000,2000\n\n" + str(point.x / 2) + ",0\n-8000,-5000\n")

    def move_side_view(self):
        """
        This method moves the side view for properly location inside the blueprint, via AutoCAD console;
        """
        self.app.ActiveDocument.SendCommand("m\n10000,-2000\n-1000,2000\n\n0,0\n5000,0\n")

    def change_scale(self, scale):
        """
        This method changes the current scale inside the AutoCAD model view, via AutoCAD console;
        :param scale: desired scale;
        :return: the desired scale is set inside the AutoCAD model view.
        """
        return self.app.ActiveDocument.SendCommand("CANNOSCALE\n1:" + str(scale) + "\n")

    def squeeze_left(self, point_1, point_2):
        """
        This method implies the "stretch" AutoCAD command from one to another of specified points on the left side,
        via AutoCAD console;
        :param point_1: the starting point;
        :param point_2: the endpoint.
        """
        self.app.ActiveDocument.SendCommand("_stretch\n" + str(point_1.x + 2500) + ",-2000\n" +
                                            str(point_1.x - 2500) + ",2000\n\n0,0\n" +
                                            str((point_2.x - 4500)/2) + ",0\n")

    def squeeze_right(self, point_1, point_2):
        """
        This method implies the "stretch" AutoCAD command from one to another of specified points on the right side,
        via AutoCAD console;
        :param point_1: the starting point;
        :param point_2: the endpoint.
        """
        self.app.ActiveDocument.SendCommand("_stretch\n" + str(point_1.x + 2500) + ",-2000\n" +
                                            str(point_1.x - 2500) + ",2000\n\n0,0\n" +
                                            str(-(point_2.x - 4500)/2) + ",0\n")

    def original_element(self, offset):
        """
        This method copies all of the created views which don't have yet any design elements
        to the top right corner of the current blueprint, via AutoCAD console;
        :param offset: desired offset.
        """
        self.app.ActiveDocument.SendCommand("c\n25000,-2000\n-1000,2000\n\n0,0\n" + str(offset) + "\n\n")
