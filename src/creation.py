from src.classes import *


def create_front_view(app, worksheet, scale, holes_offset, g):
    edit = Editor(app)
    edit.change_scale(scale)

    constr = Constructor(app)

    constr.add_polyline(g['FGP'], 0, 'Основная 0.25')
    constr.add_polyline(g['FGP'], 1, 'Основная 0.25')
    constr.add_polyline(g['FGP'], 2, 'Основная 0.25')

    constr.add_block_group(g['LWH'], 'hole_25')
    constr.add_block_group(g['RWH'], 'hole_25')

    # bottom line
    dim = Dimensions(app)
    dim.rotated_dim_x(g['FGP'].get_point(0, 0).center,
                      g['LWH'].get_point(0, 0).offset(0, -holes_offset),
                      g['FGP'].get_point(0, 0), -10, scale)
    dim.chain_rotated_dim_x(g['LWH'], g['FGP'].get_point(0, 0), -holes_offset, -10, scale)
    dim.rotated_dim_x(g['LWH'].get_point(len(g['LWH'].rows)-1, 0).offset(0, -holes_offset),
                      g['RWH'].get_point(len(g['RWH'].rows)-1, 0).offset(0, -holes_offset),
                      g['FGP'].get_point(0, 0), -10, scale)
    dim.chain_rotated_dim_x(g['RWH'], g['FGP'].get_point(0, 0), -holes_offset, -10, scale)
    dim.rotated_dim_x(g['RWH'].get_point(0, 0).offset(0, -holes_offset),
                      g['FGP'].get_point(0, 1).center,
                      g['FGP'].get_point(0, 0), -10, scale)

    # left line
    dim.rotated_dim_y(g['FGP'].get_point(0, 0).center,
                      g['LWH'].get_point(0, 0).offset(-holes_offset, 0),
                      g['FGP'].get_point(0, 0), -10, scale)
    # ======================================
    dim.rotated_dim_y(g['LWH'].get_point(0, 0).offset(-holes_offset, 0),
                      g['LWH'].get_upper_right().offset(-holes_offset, 0),
                      g['FGP'].get_point(0, 0), -10, scale).\
        TextOverride = Editor(app).override_dim(g['LWH'], [0, 0], [0, g['LWH'].row_len(0)-1])
    # ======================================
    dim.rotated_dim_y(g['FGP'].get_point(1, 3).center,
                      g['LWH'].get_upper_right().offset(-holes_offset, 0),
                      g['FGP'].get_point(0, 0), -10, scale)

    # right line
    dim.rotated_dim_y(g['FGP'].get_point(0, 1).center,
                      g['RWH'].get_point(0, 0).offset(holes_offset, 0),
                      g['FGP'].get_point(0, 1), 10, scale)
    dim.rotated_dim_y(g['RWH'].get_point(0, 0).offset(holes_offset, 0),
                      g['RWH'].get_upper_right().offset(holes_offset, 0),
                      g['FGP'].get_point(0, 1), 10, scale).\
        TextOverride = edit.override_dim(g['LWH'], [0, 0], [0, g['LWH'].row_len(0)-1])

    dim.rotated_dim_y(g['FGP'].get_point(1, 2).center,
                      g['RWH'].get_upper_right().offset(holes_offset, 0),
                      g['FGP'].get_point(0, 1), 10, scale)

    # top line
    dim.rotated_dim_x(g['FGP'].get_point(1, 3).center,
                      g['FGP'].get_point(1, 2).center,
                      g['FGP'].get_point(1, 2), 10, scale)

    design = Design(app, worksheet)
    design.main_header(scale, g['FGP'])

    holes = Leader(app, worksheet)
    holes.add_mleader_holes(g['LWH'].get_down_left().center,
                            g['FGP'].get_upper_left().
                            offset(g['LWH'].get_down_left().x + 5 * scale, -5 * scale),
                            g['LWH'])
    holes.add_mleader_holes(g['RWH'].get_upper_right().center,
                            g['FGP'].get_point(1, 2).
                            offset(5 * scale, 10 * scale),
                            g['RWH'])
    leader = Leader(app, worksheet)
    leader.add_mleader_pos(g['FGP'].get_point(1, 3).offset(1000, 0),
                           g['FGP'].get_point(1, 3).offset(1000, 13 * scale), '2')
    if g['TGP'].get_point(0, 1).center != g['TGP'].get_point(1, 1).center:
        text_pos1 = '3'
    else:
        text_pos1 = '2'
    leader.add_mleader_pos(g['FGP'].get_upper_left().offset(1000, 0),
                           g['FGP'].get_upper_left().offset(1000, -17 * scale), text_pos1)

    leader.add_mleader_pos(g['FGP'].get_point(0, 1).offset(0, 80),
                           g['FGP'].get_point(0, 1).offset(5 * scale, -10 * scale), '1')
    design.insert_block(g['FGP'].get_point(1, 3).offset(5, 15 * scale), 'sec_1_t', scale)
    design.insert_block(g['FGP'].get_upper_left().offset(5, -15 * scale), 'sec_1_b', scale)
    design.insert_block(g['FGP'].get_point(1, 2).
                        offset(float(-g['FGP'].get_point(1, 2).x / 2), 15 * scale),
                        'view_A_t', scale)
    design.insert_block(g['FGP'].get_point(0, 1).
                        offset(float(-g['FGP'].get_point(0, 1).x / 2), -15 * scale),
                        'view_po_A_b', scale)
    edit.move_front_view()


def create_top_view(app, worksheet, scale, holes_offset, g):
    edit = Editor(app)
    edit.change_scale(scale)

    constr = Constructor(app)

    constr.add_block_group(g['LCH'], 'hole_25')
    constr.add_block_group(g['MCH'], 'hole_25')
    constr.add_block_group(g['RCH'], 'hole_25')
    constr.add_polyline(g['TGP'], 0, 'Основная 0.25')
    if g['TGP'].get_point(0, 1).center != g['TGP'].get_point(1, 1).center:
        constr.add_polyline(g['TGP'], 1, 'Основная 0.25')
    constr.add_polyline(g['TGP'], 2, 'Ось')
    constr.add_polyline(g['TGP'], 3, 'Пунктир')

    dim = Dimensions(app)
    # bottom line
    dim.rotated_dim_x(g['TGP'].get_point(0, 0).center,
                      g['LCH'].get_point(0, 0).offset(0, -holes_offset),
                      g['TGP'].get_point(0, 0), -10, scale)
    dim.chain_rotated_dim_x(g['LCH'], g['TGP'].get_point(0, 0), -holes_offset, -10, scale)
    if len(g['MCH'].rows) > 0:
        dim.rotated_dim_x(g['LCH'].get_point(len(g['LCH'].rows) - 1, 0).offset(0, -holes_offset),
                          g['MCH'].get_point(0, 0).offset(0, -holes_offset),
                          g['TGP'].get_point(0, 0), -10, scale)

        dim.chain_rotated_dim_x(g['MCH'], g['TGP'].get_point(0, 0), -holes_offset, -10, scale)
        dim.rotated_dim_x(g['MCH'].get_point(len(g['MCH'].rows)-1, 0).offset(0, -holes_offset),
                          g['RCH'].get_point(len(g['RCH'].rows)-1, 0).offset(0, -holes_offset),
                          g['TGP'].get_point(0, 0), -10, scale)
    else:
        dim.rotated_dim_x(g['LCH'].get_point(len(g['LCH'].rows) - 1, 0).offset(0, -holes_offset),
                          g['RCH'].get_point(len(g['RCH'].rows) - 1, 0).offset(0, -holes_offset),
                          g['TGP'].get_point(0, 0), -10, scale)
    dim.chain_rotated_dim_x(g['RCH'], g['TGP'].get_point(0, 0), -holes_offset, -10, scale)
    dim.rotated_dim_x(g['RCH'].get_point(0, 0).offset(0, -holes_offset),
                      g['TGP'].get_point(0, 1).center,
                      g['TGP'].get_point(0, 0), -10, scale)

    # left line
    dim.rotated_dim_y(g['TGP'].get_point(0, 0).center,
                      g['LCH'].get_point(0, 0).offset(-holes_offset, 0),
                      g['TGP'].get_point(0, 0), -10, scale)
    dim.chain_rotated_dim_y(g['LCH'], g['TGP'].get_point(0, 0), -holes_offset, -10, scale)
    dim.rotated_dim_y(g['TGP'].get_upper_right().center,
                      g['LCH'].get_upper_right().offset(-holes_offset, 0),
                      g['TGP'].get_point(0, 0), -10, scale)

    # right line
    dim.rotated_dim_y(g['TGP'].get_point(0, 1).center,
                      g['RCH'].get_point(0, 0).offset(holes_offset, 0),
                      g['TGP'].get_point(0, 1), 10, scale)
    dim.chain_rotated_dim_y(g['RCH'], g['TGP'].get_point(0, 1), holes_offset, 10, scale)
    dim.rotated_dim_y(g['TGP'].get_point(0, 2).center,
                      g['RCH'].get_upper_right().offset(holes_offset, 0),
                      g['TGP'].get_point(0, 1), 10, scale)

    # top line
    dim.rotated_dim_x(g['TGP'].get_point(0, 3).center,
                      g['TGP'].get_point(0, 2).center,
                      g['TGP'].get_point(0, 2), 10, scale)

    # center line
    if g['MCH'].grid_len() > 0:
        dim.rotated_dim_y(g['TGP'].get_point(0, 0).
                          offset(g['MCH'].get_upper_left().x - 200, 0),
                          g['MCH'].get_point(0, 0).offset(-holes_offset, 0),
                          g['MCH'].get_point(0, 0), -10, scale)
        dim.chain_rotated_dim_y(g['MCH'], g['MCH'].get_point(0, 0), -holes_offset, -10, scale)
        dim.rotated_dim_y(g['TGP'].get_upper_right().
                          offset(g['MCH'].get_upper_left().x - 200, 0),
                          g['MCH'].get_upper_right().offset(-holes_offset, 0),
                          g['MCH'].get_point(0, 0), -10, scale)
        holes = Leader(app, worksheet)
        holes.add_mleader_holes(g['MCH'].get_down_right().center,
                                g['TGP'].get_point(0, 3).
                                offset(g['MCH'].get_down_right().x + 5 * scale, 5 * scale), g['MCH'])

    design = Design(app, worksheet)
    design.header(u'A (1:' + str(scale) + ')', g['TGP'], scale)
    holes = Leader(app, worksheet)
    holes.add_mleader_holes(g['LCH'].get_point(len(g['LCH'].rows) - 1, 0).center,
                            g['TGP'].get_point(0, 0).
                            offset(g['LCH'].get_point(len(g['LCH'].rows) - 1, 0).x +
                                   5 * scale, - 5 * scale), g['LCH'])

    holes.add_mleader_holes(g['RCH'].get_upper_right().center,
                            g['TGP'].get_point(0, 2).
                            offset(5 * scale, 10 * scale), g['RCH'])
    if g['TGP'].get_point(0, 1).center != g['TGP'].get_point(1, 1).center:
        text_pos2 = '2(3)'
    else:
        text_pos2 = '2'
    leader = Leader(app, worksheet)
    leader.add_mleader_pos(g['TGP'].get_upper_left().offset(1000, 0),
                           g['TGP'].get_upper_left().offset(1000, -17 * scale), text_pos2)
    edit.move_top_view()


def create_side_view(app, worksheet, scale, g):
    edit = Editor(app)
    edit.change_scale(scale)

    constr = Constructor(app)
    constr.add_polyline(g['SGP'], 0, 'Основная 0.25')
    constr.add_polyline(g['SGP'], 1, 'Основная 0.25')
    constr.add_polyline(g['SGP'], 2, 'Ось')
    constr.add_polyline(g['SGP'], 3, 'Основная 0.25')

    dim = Dimensions(app)
    # bottom line
    if g['SGP'].get_point(0, 1).x != g['SGP'].get_point(1, 1).x:
        dim.rotated_dim_x(g['SGP'].get_point(0, 0).center,
                          g['SGP'].get_point(0, 1).center,
                          g['SGP'].get_point(0, 0), -10, scale)

    # left line
    dim.rotated_dim_y(g['SGP'].get_point(0, 0).center,
                      g['SGP'].get_point(0, g['SGP'].row_len(0) - 1).center,
                      g['SGP'].get_point(0, 0), -10, scale)
    dim.rotated_dim_y(g['SGP'].get_point(0, g['SGP'].row_len(0) - 1).center,
                      g['SGP'].get_point(1, 0).center,
                      g['SGP'].get_point(0, 0), -10, scale)
    dim.rotated_dim_y(g['SGP'].get_point(1, 0).center,
                      g['SGP'].get_point(1, g['SGP'].row_len(1) - 1).center,
                      g['SGP'].get_point(0, 0), -10, scale)

    # right line
    dim.rotated_dim_y(g['SGP'].get_point(0, 1).center,
                      g['SGP'].get_point(1, 2).center,
                      g['SGP'].get_point(0, 1), 10, scale)

    # top line
    dim.rotated_dim_x(g['SGP'].get_point(1, 2).center,
                      g['SGP'].get_point(1, g['SGP'].row_len(0) - 1).center,
                      g['SGP'].get_point(1, 2), 10, scale)

    design = Design(app, worksheet)
    design.header(u'1-1 (1:' + str(scale) + ')', g['SGP'], scale)

    leader = Leader(app, worksheet)
    leader.add_mleader_pos(g['SGP'].get_point(1, 2).offset(-2 * scale, 0),
                           g['SGP'].get_point(1, 2).offset(7 * scale, 10 * scale), '2')
    if g['TGP'].get_point(0, 1).center != g['TGP'].get_point(1, 1).center:
        text_pos3 = '3'
    else:
        text_pos3 = '2'
    leader.add_mleader_pos(g['SGP'].get_point(0, 1).offset(-2 * scale, 0),
                           g['SGP'].get_point(0, 1).offset(7 * scale, -10 * scale), text_pos3)
    leader = Leader(app, worksheet)
    leader.add_mleader_pos(g['SGP'].get_point(3, 2).offset(0, -g['SGP'].get_point(3, 2). y / 2),
                           g['SGP'].get_point(3, 2).offset(7 * scale,
                                                           -10 * scale
                                                           - g['SGP'].get_point(3, 2). y / 2), '1')

    edit.move_side_view()
