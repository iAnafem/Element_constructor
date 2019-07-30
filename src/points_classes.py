import array


class Point:
    """This is a class that represents a single geometry point"""
    def __init__(self, coord):
        """
        The constructor for Point class.
        :param coord: a list of coordinates
        Attributes:
            center: represent the point in space. 'd' - means 'double'; [...] - a list of coordinates;
            x, y, z: a particular coordinate of the point.
        """
        self.center = array.array('d', [coord[0], coord[1], coord[2]])
        self.x = coord[0]
        self.y = coord[1]
        self.z = coord[2]

    def offset(self, offset_x=0, offset_y=0):
        """
        This method creates a new temporarily point, that is offset from the current point.
        :param offset_x: x-axis offset;
        :param offset_y: y-axis offset;
        :return: a new temporarily point (object of the class Point)
        """
        return Point([self.x + offset_x, self.y + offset_y, self.z])


class Row:
    """
    This is a class that represents a row of points
    """
    def __init__(self, points: list):
        """
        The constructor for class Row.
        :param points: a list of the Point class objects;
        """
        self.points = list()
        for i in points:
            self.points.append(Point(i))


class Grid:
    """
    This is a class that represents a grid of points.
    It consist of class Row objects.
    """
    def __init__(self):
        """
        The constructor for class Row.
        Attributes:
            rows: a list of the class Row objects.
        """
        self.rows = list()

    def row_len(self, row_num: int) -> int:
        """
        This method returns the length (the number of points) of the desired row.
        :param row_num: number of the desired row;
        :return: length of the desired row.
        """
        return len(self.rows[row_num].points)

    def grid_len(self) -> int:
        """
        This method returns the length (the number of points) of the current grid.
        :return: length of the current grid.
        """
        result = 0
        for i in range(len(self.rows)):
            result += len(self.rows[i].points)
        return result

    def get_point(self, row: int, column: int) -> Point:
        """
        This method returns the desired point at the intersection of the selected row and column.
        :param row: number of row of points in the current grid;
        :param column: number of column of points in the current grid
        :return: class Point object.
        """
        if row >= len(self.rows):
            raise ValueError('row out of range')
        if column >= len(self.rows[row].points):
            raise ValueError('column out of range')
        return self.rows[row].points[column]

    def get_upper_left(self) -> Point:
        """
        :return: class Point object from the upper left corner.
        """
        return self.get_point(0, 0)

    def get_upper_right(self) -> Point:
        """
        :return: class Point object from the upper right corner.
        """
        return self.get_point(0, len(self.rows[0].points) - 1)

    def get_down_left(self) -> Point:
        """
        :return: class Point object from the down left corner.
        """
        return self.get_point(len(self.rows) - 1, 0)

    def get_down_right(self) -> Point:
        """
        :return: class Point object from the down right corner.
        """
        last_row = len(self.rows) - 1
        return self.get_point(last_row, len(self.rows[last_row].points) - 1)
