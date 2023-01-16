from math import *
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import interpolate
from shapely import geometry, affinity

from utils.utils import change_variant, add_view, add_layer, add_text, add_rect, add_raster, iDocument2D, \
    kompas_document_2d, kompas_document, get_angle_between_points, endpoint_by_distance_and_angle, m_to_mm, mm_to_m, \
    interpolate_line, sum_tuple, get_intersections, line_len, draw_meadow, f_angle, f_slope, get_next_layer_id

ONE_TO_SCALE = 2000
WORKBOOK_PATH = Path("../../Геодезия.xlsm").absolute()


def get_points_dict(workbook_path: Path) -> dict:
    table_79 = pd.read_excel(workbook_path, sheet_name='pandasPlan79')
    table_10 = pd.read_excel(workbook_path, sheet_name='pandasPlan10')

    point_dict = {}

    for _, row in table_10.iterrows():
        if " - " in str(row["№№"]):
            origin_label, dest_label = row["№№"].split(" - ")

            if origin_label.isdigit():
                origin_label = int(origin_label)

            if dest_label.isdigit():
                dest_label = int(dest_label)

            origin = (table_79.loc[table_79['№№'] == origin_label]["Y"].values[0],
                      table_79.loc[table_79['№№'] == origin_label]["X"].values[0])

            _dest = (table_79.loc[table_79['№№'] == dest_label]["Y"].values[0],
                     table_79.loc[table_79['№№'] == dest_label]["X"].values[0])

            angle = np.rad2deg(get_angle_between_points(origin[0], origin[1], _dest[0], _dest[1]))
            _i_am_shit_at_coding = list(origin)
            _i_am_shit_at_coding.append(table_79.loc[table_79['№№'] == origin_label]["H"].values[0])
            point_dict.update({str(origin_label): _i_am_shit_at_coding})
            continue
        else:
            if not origin:
                continue
        _d = row["Отсчет по дальномеру kl, м"]
        _alpha = row["Отсчет по гориз. Кругу"]
        xx = origin[0] + (_d * cos(np.radians(angle - _alpha)))
        yy = origin[1] + (_d * sin(np.radians(angle - _alpha)))

        point_dict.update({str(row["№№"]): [xx, yy, row["Отметки реечн. точек Hр.т., м"]]})

    return point_dict


def add_point_marker(point_name: str, point_cords: tuple[float, float], point_height: float, radius: float = 1000, offset: tuple[float, float] = (15000, 7000)):
    iDocument2D.ksCircle(*point_cords, radius, 1)
    add_text(f'{point_name} ({str("{:1.2f}".format(point_height)).replace(".", ",")})',
             *map(lambda i, j: (i - j), point_cords, offset))


def main() -> None:
    for _variant in range(24, 25):
        change_variant(_variant, workbook_path=WORKBOOK_PATH)

        add_view("План", 1 / ONE_TO_SCALE)
        add_layer(get_next_layer_id(), 3, "Основные Точки")

        points_dict = get_points_dict(WORKBOOK_PATH)

        for point in (main_points_names := ["ПЗ41", "111", "112", "113", "ПЗ42"]):
            add_point_marker(point, m_to_mm(points_dict[point][:2]), points_dict[point][2] * 1000)

        add_layer(get_next_layer_id(), 3, "Побочные Точки")

        for point in points_dict:
            if point not in main_points_names:
                add_point_marker(point, m_to_mm(points_dict[point][:2]), points_dict[point][2] * 1000)

        interpolated_points = {}

        add_layer(get_next_layer_id(), 2, "Точки интерполяции пашни")
        # todo: переделать эту часть кода

        con_111 = "5,6,7,8,9,112,ПЗ41".split(",")

        center_point = np.array(points_dict["111"])
        for point in con_111:
            point = np.array(points_dict[point])

            if center_point[2] > point[2]:
                _coordinate_dif = center_point[:2] - point[:2]
                xia = [center_point[2], point[2]]
                height_dif = range(int(ceil(point[2])), int(round(center_point[2])))
            else:
                _coordinate_dif = point[:2] - center_point[:2]
                xia = [point[2], center_point[2]]
                height_dif = range(int(ceil(center_point[2])), int(round(point[2])))

            length = round(np.sqrt(np.sum(np.power(_coordinate_dif, 2))))
            yia = [0, length]

            _alpha = np.rad2deg(get_angle_between_points(center_point[0], center_point[1], point[0], point[1]))

            f = interpolate.interp1d(xia, yia)

            for i in height_dif:
                _d = f(i)
                _interpolated_point = endpoint_by_distance_and_angle(m_to_mm(tuple(center_point)), _d * 1000, _alpha)
                # xx = center_point[0] * 1000 + (_d * 1000 * cos(np.radians(_alpha)))
                # yy = center_point[1] * 1000 + (_d * 1000 * sin(np.radians(_alpha)))
                iDocument2D.ksPoint(*_interpolated_point, 0)
                add_text(str(i), _interpolated_point[0] - 10000, _interpolated_point[1] - 8000, color=9410565)

                if i in interpolated_points:
                    interpolated_points[i].append(_interpolated_point)
                else:
                    interpolated_points.update({i: [_interpolated_point]})

        x = [points_dict[p][0] for p in points_dict]
        y = [points_dict[p][1] for p in points_dict]
        centroid = (sum(x) / len(points_dict), sum(y) / len(points_dict))
        iDocument2D.ksPoint(*(cord * 1000 for cord in centroid), 0)

        current_view = kompas_document_2d.ViewsAndLayersManager.Views.ActiveView
        current_view.X = 594 / 2 - centroid[0] / 2
        current_view.Y = 420 / 2 - centroid[1] / 2
        current_view.Update()

        interpolate_line("7", "8", interpolated_points, points_dict)

        add_layer(get_next_layer_id(), 3, "Кривые интерполяции пашни")

        _max_num_of_points = max([len(interpolated_points[height_points]) for height_points in interpolated_points])

        for height in interpolated_points:

            if len(interpolated_points[height]) <= 3:
                continue

            cent = (sum([p[0] for p in interpolated_points[height]]) / len(interpolated_points[height]),
                    sum([p[1] for p in interpolated_points[height]]) / len(interpolated_points[height]))
            interpolated_points[height].sort(key=lambda p: atan2(p[1] - cent[1], p[0] - cent[0]))

            iDocument2D.ksBezier(len(interpolated_points[height]) == _max_num_of_points, 1)
            for point in interpolated_points[height]:
                iDocument2D.ksPoint(*point, 0)
            iDocument2D.ksEndObj()

        add_layer(get_next_layer_id(), 2, "Точки интерполяции берега")

        interpolated_points_shore = {}

        # ПЗ 41
        interpolate_line("4", "5", interpolated_points_shore, points_dict)
        interpolate_line("ПЗ41", "1", interpolated_points_shore, points_dict)
        interpolate_line("ПЗ41", "2", interpolated_points_shore, points_dict)
        interpolate_line("ПЗ41", "4", interpolated_points_shore, points_dict)
        interpolate_line("ПЗ41", "5", interpolated_points_shore, points_dict)
        interpolate_line("ПЗ41", "21", interpolated_points_shore, points_dict)

        # 112
        # interpolate_line("8", "10", interpolated_points_shore, extra_points)
        # interpolate_line("112", "10", interpolated_points_shore, extra_points)
        # interpolate_line("112", "13", interpolated_points_shore, extra_points)
        # interpolate_line("9", "12", interpolated_points_shore, extra_points)
        interpolate_line("12", "13", interpolated_points_shore, points_dict)
        # interpolate_line("13", "1", interpolated_points_shore, extra_points)
        interpolate_line("1", "112", interpolated_points_shore, points_dict)

        # 113
        interpolate_line("10", "14", interpolated_points_shore, points_dict)
        # interpolate_line("113", "15", interpolated_points_shore, extra_points)
        # interpolate_line("17", "113", interpolated_points_shore, extra_points)
        interpolate_line("16", "113", interpolated_points_shore, points_dict)
        # interpolate_line("17", "112", interpolated_points_shore, extra_points)
        # interpolate_line("17", "14", interpolated_points_shore, extra_points)

        # ПЗ 42
        interpolate_line("ПЗ42", "20", interpolated_points_shore, points_dict)
        interpolate_line("ПЗ42", "18", interpolated_points_shore, points_dict)
        interpolate_line("23", "22", interpolated_points_shore, points_dict)
        interpolate_line("22", "21", interpolated_points_shore, points_dict)
        interpolate_line("20", "21", interpolated_points_shore, points_dict)
        interpolate_line("1", "21", interpolated_points_shore, points_dict)
        interpolate_line("21", "21", interpolated_points_shore, points_dict)

        add_layer(get_next_layer_id(), 3, "Кривые интерполяции берега")

        for height in interpolated_points_shore:

            if len(interpolated_points_shore[height]) < 3:
                continue

            interpolated_points_shore[height].sort()

            iDocument2D.ksBezier(False, 1)
            for point in interpolated_points_shore[height]:
                iDocument2D.ksPoint(*point, 0)
            iDocument2D.ksEndObj()

        # Горизонталь вокруг ФС
        add_layer(get_next_layer_id(), 2, "Точки интерполяции ФС")

        interpolated_points_fruit_garden = {}

        _interpolation_21_23 = interpolate_line("21", "23", {}, points_dict)
        _last_point_between_21_23 = list(_interpolation_21_23.keys())[-1]

        interpolated_points_fruit_garden.update(
            {_last_point_between_21_23: _interpolation_21_23[_last_point_between_21_23]})

        _interpolation_112_17 = interpolate_line("112", "17", {}, points_dict)
        if _last_point_between_21_23 in _interpolation_112_17:
            interpolated_points_fruit_garden[_last_point_between_21_23].append(
                *_interpolation_112_17[_last_point_between_21_23])

        _interpolation_14_17 = interpolate_line("14", "17", {}, points_dict)
        interpolated_points_fruit_garden[_last_point_between_21_23].append(
            *_interpolation_14_17[_last_point_between_21_23])

        _interpolation_15_113 = interpolate_line("15", "113", {}, points_dict)
        interpolated_points_fruit_garden[_last_point_between_21_23].append(
            *_interpolation_15_113[_last_point_between_21_23])

        _interpolation_p42_18 = interpolate_line("ПЗ42", "18", {}, points_dict)
        interpolated_points_fruit_garden[_last_point_between_21_23].insert(0, *_interpolation_p42_18[
            _last_point_between_21_23])

        add_layer(get_next_layer_id(), 3, "Кривые интерполяции ФС")

        iDocument2D.ksBezier(False, 1)
        for point in interpolated_points_fruit_garden[_last_point_between_21_23]:
            iDocument2D.ksPoint(*point, 0)
        iDocument2D.ksEndObj()

        # Горизонталь за ФС
        add_layer(get_next_layer_id(), 2, "Точки интерполяции за ФС")

        interpolated_points_behind_fruit_garden = {}

        _interpolation_112_10 = interpolate_line("112", "10", {}, points_dict)
        _last_point_between_112_10 = list(_interpolation_112_10.keys())[-1]

        interpolated_points_behind_fruit_garden.update(
            {_last_point_between_112_10: _interpolation_112_10[_last_point_between_112_10]})

        _interpolation_14_17 = interpolate_line("14", "17", {}, points_dict)
        interpolated_points_behind_fruit_garden[_last_point_between_112_10].append(
            *_interpolation_14_17[_last_point_between_112_10])

        _interpolation_15_113 = interpolate_line("15", "113", {}, points_dict)
        if _last_point_between_112_10 in _interpolation_15_113:
            interpolated_points_behind_fruit_garden[_last_point_between_112_10].append(
                *_interpolation_15_113[_last_point_between_112_10])

        _interpolation_8_10 = interpolate_line("8", "10", {}, points_dict)
        interpolated_points_behind_fruit_garden[_last_point_between_112_10].insert(0, *_interpolation_8_10[
            _last_point_between_112_10])

        add_layer(get_next_layer_id(), 3, "Кривые интерполяции за ФС")

        iDocument2D.ksBezier(False, 1)
        for point in interpolated_points_behind_fruit_garden[_last_point_between_112_10]:
            iDocument2D.ksPoint(*point, 0)
        iDocument2D.ksEndObj()

        # Соединить основные границы

        add_layer(get_next_layer_id(), 3, "Левая граница")

        _left_border = "4,5,6,7".split(sep=",")

        for i in range(len(_left_border) - 1):
            # print(f"{ _left_border[i]}: {extra_points[ _left_border[i]]}\t"
            #       f"{ _left_border[i+1]}: {extra_points[ _left_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*points_dict[_left_border[i]][:2], *points_dict[_left_border[i + 1]][:2])], 4)

        add_layer(get_next_layer_id(), 3, "Верхняя граница")

        _top_border = "7,8,10,14,15".split(sep=",")

        for i in range(len(_top_border) - 1):
            # print(f"{ _top_border[i]}: {extra_points[ _top_border[i]]}\t"
            #       f"{ _top_border[i+1]}: {extra_points[ _top_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*points_dict[_top_border[i]][:2], *points_dict[_top_border[i + 1]][:2])], 1)

        add_layer(get_next_layer_id(), 3, "Граница пашни")

        _p_border = "5,ПЗ41,9,112,8".split(sep=",")

        for i in range(len(_p_border) - 1):
            # print(f"{ _p_border[i]}: {extra_points[ _p_border[i]]}\t"
            #       f"{ _p_border[i+1]}: {extra_points[ _p_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*points_dict[_p_border[i]][:2], *points_dict[_p_border[i + 1]][:2])], 4)

        add_layer(get_next_layer_id(), 3, "Река")

        _river_border = "4,2,21,20".split(sep=",")

        for i in range(len(_river_border) - 1):
            # print(f"{ _river_border[i]}: {extra_points[ _river_border[i]]}\t"
            #       f"{ _river_border[i+1]}: {extra_points[ _river_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*points_dict[_river_border[i]][:2], *points_dict[_river_border[i + 1]][:2])], 1)

        # вторая сторона реки

        _alpha = np.rad2deg(get_angle_between_points(*points_dict["4"][:2], *points_dict["3"][:2]))

        _second_shore = []
        _d = 32
        for i in range(len(_river_border)):
            _second_shore.append(endpoint_by_distance_and_angle(
                m_to_mm(points_dict[_river_border[i]][:2]), _d * 1000, _alpha))

        for i in range(len(_second_shore) - 1):
            if i != len(_second_shore) - 2:
                iDocument2D.ksLineSeg(*_second_shore[i], *_second_shore[i + 1], 1)
            else:
                iDocument2D.ksLineSeg(*_second_shore[i], *(p * 1000 for p in points_dict["19"][:2]), 1)

        # Добавить подпись

        _xx = (points_dict["2"][0] + points_dict["21"][0]) / 2 * 1000 + (_d / 2 * 1000 * cos(np.radians(_alpha)))
        _yy = (points_dict["2"][1] + points_dict["21"][1]) / 2 * 1000 + (_d / 2 * 1000 * sin(np.radians(_alpha)))

        add_text("р. Соть", _xx, _yy,
                 np.rad2deg(get_angle_between_points(*points_dict["2"][:2], *points_dict["21"][:2])), 5)

        # Автодорога

        add_layer(get_next_layer_id(), 3, "Автодорога")

        _alpha = np.rad2deg(get_angle_between_points(*points_dict["6"][:2], *points_dict["7"][:2]))

        _autobahn = []
        _d = 10
        for i in range(len(_top_border)):
            xx = points_dict[_top_border[i]][0] * 1000 + (_d * 1000 * cos(np.radians(_alpha)))
            yy = points_dict[_top_border[i]][1] * 1000 + (_d * 1000 * sin(np.radians(_alpha)))
            _autobahn.append((xx, yy))

        for i in range(len(_autobahn) - 1):
            iDocument2D.ksLineSeg(*_autobahn[i], *_autobahn[i + 1], 4)

        # Текст пашни
        add_layer(get_next_layer_id(), 3, "Текст пашни")

        add_text("Пашня", *m_to_mm(sum_tuple(points_dict["111"][:2], (-20, 45))), 0, 7)

        # Текст лес
        add_layer(get_next_layer_id(), 3, "Текст лес")

        add_text("Лес", *m_to_mm(sum_tuple(points_dict["6"][:2], (-60, 0))), 0, 7)

        # Фруктовый сад
        add_layer(get_next_layer_id(), 3, "Колодец")

        _alpha_14_15 = np.rad2deg(get_angle_between_points(*points_dict["14"][:2], *points_dict["15"][:2]))
        _alpha_113_P42 = np.rad2deg(get_angle_between_points(*points_dict["113"][:2], *points_dict["ПЗ42"][:2]))

        # _well_angles = degrees(atan2(14.62,9.15))
        #
        _well_point_1 = m_to_mm(endpoint_by_distance_and_angle(points_dict["113"][:2], 40, _alpha_113_P42))
        # _obj_well_point_1 = iDocument2D.ksPoint(*_well_point_1, 0)

        _well_point_2 = m_to_mm(endpoint_by_distance_and_angle(points_dict["113"][:2], 58.61, _alpha_113_P42))
        # _obj_well_point_2 = iDocument2D.ksPoint(*_well_point_2, 0)
        #
        # well_line_1 = iDocument2D.ksLineSeg(*_well_point_1, *m_to_mm(endpoint_by_distance_and_angle(mm_to_m(
        # _well_point_1), 9.15, _alpha_113_P42 - _well_angles)), 6)

        _well_point_3 = get_intersections(*_well_point_1, 9.15 * 1000, *_well_point_2, 14.62 * 1000)

        # Штриховка
        _well_hatch = iDocument2D.ksHatch(0, 45, 0.25, 0, 0, 0)
        _obj_well_circle = iDocument2D.ksCircle(*_well_point_3[:2], 2500, 1)
        iDocument2D.ksEndObj()

        _obj_well_circle = iDocument2D.ksCircle(*_well_point_3[:2], 2500, 1)

        add_text("Колодец", *m_to_mm(sum_tuple(mm_to_m(_well_point_3[:2]), (-10, 5))))

        add_layer(get_next_layer_id(), 3, "Фруктовый сад")

        # fruit_garden_points = [endpoint_by_distance_and_angle(extra_points["113"][:2], 81.5, _alpha_113_P42 - (44 +
        # (1/60))), endpoint_by_distance_and_angle(endpoint_by_distance_and_angle(extra_points["113"][:2], 58.61,
        # _alpha_113_P42), -8.2, _alpha_14_15), endpoint_by_distance_and_angle(endpoint_by_distance_and_angle(
        # extra_points["113"][:2], 133.41, _alpha_113_P42), -7.81, _alpha_14_15), endpoint_by_distance_and_angle(
        # extra_points["ПЗ42"][:2], 96.15, 0 - _alpha_113_P42 - (25 + (11/60)))]

        fruit_garden_points = [tuple(points_dict["17"][:2]), endpoint_by_distance_and_angle(
            endpoint_by_distance_and_angle(points_dict["113"][:2], 58.61, _alpha_113_P42), -8.2, _alpha_14_15),
                               endpoint_by_distance_and_angle(
                                   endpoint_by_distance_and_angle(points_dict["113"][:2], 133.41, _alpha_113_P42),
                                   -7.81,
                                   _alpha_14_15), tuple(points_dict["23"][:2])]

        iDocument2D.ksLineSeg(*m_to_mm(fruit_garden_points[0]), *m_to_mm(fruit_garden_points[3]), 1)
        for i in range(len(fruit_garden_points) - 1):
            iDocument2D.ksLineSeg(*m_to_mm(fruit_garden_points[i]), *m_to_mm(fruit_garden_points[i + 1]), 1)

        # _mid_fs_text = tuple(i/2 for i in (fruit_garden_points[0][0] + fruit_garden_points[1][0],
        # fruit_garden_points[0][1] + fruit_garden_points[1][1])) *map(lambda i, j: (i + j) * 1000, _mid_fs_text,
        # (-60, 0))

        _alpha_fs = np.rad2deg(get_angle_between_points(*fruit_garden_points[1], *fruit_garden_points[2]))

        add_text("ФС", *m_to_mm(endpoint_by_distance_and_angle(fruit_garden_points[0], 15, _alpha_fs + 30)),
                 np.rad2deg(get_angle_between_points(*fruit_garden_points[0], *fruit_garden_points[1])), 5)

        add_layer(get_next_layer_id(), 3, "2КЖ")

        _kg2_p1 = m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(points_dict["113"][:2], 80.05, _alpha_113_P42),
                -8.03, _alpha_14_15))
        _kg2_p2 = m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(points_dict["113"][:2], 110.23, _alpha_113_P42),
                -7.91, _alpha_14_15))

        add_rect(_kg2_p1, -16.05 * 1000, (110.23 - 80.05) * 1000, _alpha_fs)

        # *tuple(i/2 for i in (_kg2_p1[0] + _kg2_p2[0], _kg2_p1[1] + _kg2_p2[1]))

        add_text("2кж", _kg2_p2[0] - 2400, _kg2_p2[1] + 4000, np.rad2deg(get_angle_between_points(*_kg2_p2, *_kg2_p1)),
                 5)

        # Железная дорога
        add_layer(get_next_layer_id(), 3, "Железная дорога")

        _railroad_p1 = endpoint_by_distance_and_angle(points_dict["113"][:2], 12.64, _alpha_14_15)
        _railroad_end = endpoint_by_distance_and_angle(
            endpoint_by_distance_and_angle(points_dict["113"][:2], 162.1, _alpha_113_P42), 14.28, _alpha_14_15)

        _alpha_rail = np.rad2deg(get_angle_between_points(*_railroad_p1, *_railroad_end))

        _railroad_p0 = endpoint_by_distance_and_angle(_railroad_p1, -70, _alpha_rail)

        iDocument2D.ksLineSeg(*m_to_mm(_railroad_p0), *m_to_mm(_railroad_end), 1)

        for i in range(5, int(line_len(_railroad_p0, _railroad_end)) + 5, 5):
            _start_point = endpoint_by_distance_and_angle(_railroad_p0, i, _alpha_rail)
            _side_point_1 = endpoint_by_distance_and_angle(_start_point, 2.5, 45 - _alpha_rail)
            _side_point_2 = endpoint_by_distance_and_angle(_start_point, -2.5, 45 - _alpha_rail)
            iDocument2D.ksLineSeg(*m_to_mm(_side_point_1), *m_to_mm(_side_point_2), 1)

        # Условные обозначения леса, луга, фруктового сада.

        _plan_poly = [m_to_mm(tuple(points_dict["20"][:2])), m_to_mm(tuple(points_dict["21"][:2])),
                      m_to_mm(tuple(points_dict["2"][:2])),
                      m_to_mm(tuple(points_dict["4"][:2])), m_to_mm(tuple(points_dict["5"][:2])),
                      m_to_mm(tuple(points_dict["6"][:2])),
                      m_to_mm(tuple(points_dict["7"][:2])), m_to_mm(tuple(points_dict["8"][:2])),
                      m_to_mm(tuple(points_dict["10"][:2])),
                      m_to_mm(tuple(points_dict["14"][:2])), m_to_mm(tuple(points_dict["15"][:2])),
                      m_to_mm(tuple(points_dict["113"][:2])),
                      m_to_mm(tuple(points_dict["ПЗ42"][:2]))]

        _plan_poly_line = geometry.LineString(_plan_poly)
        _plan_poly = geometry.Polygon(_plan_poly_line)

        _farm_poly = [m_to_mm(tuple(points_dict["ПЗ41"][:2])), m_to_mm(tuple(points_dict["5"][:2])),
                      m_to_mm(tuple(points_dict["6"][:2])),
                      m_to_mm(tuple(points_dict["7"][:2])), m_to_mm(tuple(points_dict["8"][:2])),
                      m_to_mm(tuple(points_dict["112"][:2])),
                      m_to_mm(tuple(points_dict["9"][:2]))]

        _farm_poly_line = geometry.LineString(_farm_poly)
        _farm_poly = geometry.Polygon(_farm_poly_line)

        _fruit_garden_poly = geometry.Polygon(geometry.LineString((m_to_mm(i) for i in fruit_garden_points)))

        _fruit_garden_poly = affinity.scale(_fruit_garden_poly, 0.95, 0.95)

        _kg_poly = [_kg2_p1, m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(points_dict["113"][:2], 80.05, _alpha_113_P42),
                -8.03 - 16.05, _alpha_14_15)), m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(points_dict["113"][:2], 110.23, _alpha_113_P42),
                -7.91 - 16.05, _alpha_14_15)), _kg2_p2]

        _kg_poly_line = geometry.LineString(_kg_poly)
        _kg_poly = geometry.Polygon(_kg_poly_line)

        add_layer(get_next_layer_id(), 0, "Условные обозначения луга")

        for i in range(int(points_dict["20"][1] - int(points_dict["20"][1]) % 5), int(points_dict["7"][1]), 20):
            for b in range(int(points_dict["4"][0] - int(points_dict["4"][1]) % 5), int(points_dict["15"][0]), 20):
                _d_point = geometry.Point(m_to_mm((b, i)))
                if _plan_poly.contains(_d_point) and \
                        not _farm_poly.contains(_d_point) and \
                        not _fruit_garden_poly.contains(_d_point):
                    draw_meadow(m_to_mm((b, i)))

        add_layer(get_next_layer_id(), 0, "Условные обозначения фруктового сада")

        for i in range(int(fruit_garden_points[2][1]), int(fruit_garden_points[0][1]), 5):
            for b in range(int(fruit_garden_points[-1][0]), int(fruit_garden_points[1][0]), 5):
                _d_point = geometry.Point(m_to_mm((b, i)))
                if _fruit_garden_poly.contains(_d_point) and not _kg_poly.contains(_d_point):
                    iDocument2D.ksCircle(*m_to_mm((b, i)), 1000, 2)

        add_layer(get_next_layer_id(), 0, "Условные обозначения леса")

        for i in range(int(points_dict["4"][1] - int(points_dict["4"][1]) % 5), int(points_dict["7"][1]), 20):
            for b in range(int(points_dict["3"][0] - int(points_dict["3"][1]) % 5), int(points_dict["7"][0]), 20):
                _d_point = geometry.Point(m_to_mm((b, i)))
                if not _plan_poly.contains(_d_point):
                    iDocument2D.ksCircle(*m_to_mm((b, i)), 2500, 2)

        # Дерево
        add_layer(get_next_layer_id(), 0, "Дерево")

        _a1 = radians(49 + 15 / 60)
        _a2 = radians(f_angle(f_slope(*points_dict["113"][:2], *points_dict["112"][:2]),
                              f_slope(*points_dict["ПЗ42"][:2], *points_dict["113"][:2])) - (55 + 6 / 60))

        s = line_len(points_dict["113"][:2], points_dict["112"][:2])

        _a3 = pi - _a2 - _a1

        _l1 = (s / sin(_a3)) * sin(_a1)
        _l2 = (s / sin(_a3)) * sin(_a2)

        _tree = endpoint_by_distance_and_angle(points_dict["113"][:2], _l1, degrees(_a2) + np.rad2deg(
            get_angle_between_points(*points_dict["113"][:2], *points_dict["112"][:2])))

        _tree_point = iDocument2D.ksPoint(*m_to_mm(_tree), 0)

        add_raster('tree-64x64.png', (_tree[0] / 2 - 1.847273, _tree[1] / 2), 0.025)

        # Рамка
        add_layer(get_next_layer_id(), 0, "Рамка")

        _padding = 30

        _top_left_frame_corner = m_to_mm((points_dict["3"][0] - _padding, points_dict["7"][1] + _padding))
        _bottom_right_frame_corner = m_to_mm((points_dict["15"][0] + _padding, points_dict["19"][1] - _padding))

        _wh_dif = (-1 * (_top_left_frame_corner[1] - _bottom_right_frame_corner[1]),
                   -1 * (_top_left_frame_corner[0] - _bottom_right_frame_corner[0]))
        _max_dim = max((abs(i) for i in _wh_dif))
        _wh_dif = (-1 * _max_dim, _max_dim)

        # iDocument2D.ksPoint(*_top_left_frame_corner, 0)
        # iDocument2D.ksPoint(*_bottom_right_frame_corner, 0)

        add_rect(_top_left_frame_corner, *_wh_dif)

        _padding += _padding
        _top_left_frame_corner = m_to_mm((points_dict["3"][0] - _padding, points_dict["7"][1] + _padding))
        _bottom_right_frame_corner = m_to_mm((points_dict["15"][0] + _padding, points_dict["19"][1] - _padding))

        _wh_dif = (-1 * (_top_left_frame_corner[1] - _bottom_right_frame_corner[1]),
                   -1 * (_top_left_frame_corner[0] - _bottom_right_frame_corner[0]))
        _max_dim = max((abs(i) for i in _wh_dif))
        _wh_dif = (-1 * _max_dim, _max_dim)

        add_rect(_top_left_frame_corner, *_wh_dif)

        _main_points = [tuple(int(i // 100) / 10 for i in points_dict[point][:2]) for point in
                        "ПЗ41,111,112,113,ПЗ42".split(sep=",")]

        _min_x = min(_main_points, key=lambda x: x[0])[0]
        _min_y = min(_main_points, key=lambda y: y[1])[1]

        _x_list = [round(_min_x + 0.2 * i, 1) for i in range(3)]
        _y_list = [round(_min_y + 0.2 * i, 1) for i in range(3)]

        for x_cord in _x_list:
            for y_cord in _y_list:
                iDocument2D.ksPoint(*m_to_mm(m_to_mm((x_cord, y_cord))), 8)

        for cord in _x_list:
            add_text(str(cord), int(cord * 1000000) + 2.5 * 2000, _bottom_right_frame_corner[1] + _padding / 7 * 1000,
                     90, 5)
            add_text(str(cord), int(cord * 1000000) + 2.5 * 2000, _top_left_frame_corner[1] - _padding / 2.5 * 1000, 90,
                     5)

        for cord in _y_list:
            add_text(str(cord), _top_left_frame_corner[0] + _padding / 7 * 1000, int(cord * 1000000) - 2.5 * 2000, 0, 5)
            add_text(str(cord), _bottom_right_frame_corner[0] + _padding / 2 * 1000, int(cord * 1000000) - 2.5 * 2000,
                     0, 5)

        # Добавить подпись

        add_text(f"План Масштаб 1:{ONE_TO_SCALE} Вариант: {_variant}", _top_left_frame_corner[0] + 10000,
                 _bottom_right_frame_corner[1] - 40000, 0, 14)

        if Path('../../../Watermark.png').exists():
            # Добавить картинку
            add_raster('../../../Watermark.png', (points_dict["3"][0] / 2, points_dict["20"][1] / 2), 0.4)

        kompas_document.SaveAs(str(Path(f'../{_variant}.pdf').absolute()))


if __name__ == "__main__":
    main()
