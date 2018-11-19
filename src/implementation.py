from src.creation import *
from src.settings import Settings

settings = Settings()

acad = get_autocad(settings.autocad)
sheet = get_excel(settings.excel)['P-3']

points = {name: Grid() for name in settings.groups}

for i in settings.groups:
    get_coordinates(points[i], settings.groups[i], sheet)

scale_front = settings.scale_front
scale_top = settings.scale_top
scale_side = settings.scale_side

holes_offset = settings.holes_offset

create_front_view(acad, sheet, scale_front, holes_offset, points)
create_top_view(acad, sheet, scale_top, holes_offset, points)
create_side_view(acad, sheet, scale_side, points)

table = Tables(acad, sheet)
table.add_specification()
