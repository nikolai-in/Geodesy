from math import *
from pathlib import Path
from typing import List, Tuple
from typing import Union

import LDefin2D  # Компас-3д
import MiscellaneousHelpers as miscHelpers  # Компас-3д
import numpy as np
import pandas as pd
import pythoncom  # Компас-3д
import xlwings as xw
from scipy import interpolate
from shapely import geometry, affinity
from win32com.client import Dispatch, gencache  # Компас-3д

# %%
#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                             pythoncom.IID_IDispatch))
miscHelpers.iKompasObject = kompas_object

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                             pythoncom.IID_IDispatch))
miscHelpers.iApplication = application

documents = application.Documents

#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()


def add_layer(number: int, state: int = 3, name: str = None) -> None:
    obj = iDocument2D.ksLayer(number)
    i_layer_param = kompas6_api5_module.ksLayerParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_LayerParam))
    i_layer_param.Init()
    if name:
        i_layer_param.name = name
    i_layer_param.state = state
    iDocument2D.ksSetObjParam(obj, i_layer_param, LDefin2D.ALLPARAM)


def add_view(name: str, scale: float = 1, x: float = 0, y: float = 0) -> None:
    i_view_param = kompas6_api5_module.ksViewParam(kompas_object.GetParamStruct(kompas6_constants.ko_ViewParam))
    i_view_param.Init()
    i_view_param.angle = 0
    i_view_param.color = 0
    i_view_param.name = name
    i_view_param.scale_ = scale
    i_view_param.state = 3
    i_view_param.x = x
    i_view_param.y = y
    iDocument2D.ksCreateSheetView(i_view_param, 0)


def add_text(text: str, x: float, y: float, angle: float = 0, font_size: float = 2.5, color=0) -> None:
    i_paragraph_param = kompas6_api5_module.ksParagraphParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_ParagraphParam))
    i_paragraph_param.Init()
    i_paragraph_param.x = x
    i_paragraph_param.y = y
    i_paragraph_param.ang = angle
    i_paragraph_param.height = 3.55
    i_paragraph_param.width = 4
    i_paragraph_param.hFormat = 0
    i_paragraph_param.vFormat = 0
    i_paragraph_param.style = 1
    iDocument2D.ksParagraph(i_paragraph_param)

    i_text_line_param = kompas6_api5_module.ksTextLineParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
    i_text_line_param.Init()
    i_text_line_param.style = 1
    i_text_item_array = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
    i_text_item_param = kompas6_api5_module.ksTextItemParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
    i_text_item_param.Init()
    i_text_item_param.iSNumb = 0
    i_text_item_param.s = text
    i_text_item_param.type = 0
    i_text_item_font = kompas6_api5_module.ksTextItemFont(i_text_item_param.GetItemFont())
    i_text_item_font.Init()
    i_text_item_font.bitVector = 4096
    i_text_item_font.color = color
    i_text_item_font.fontName = "GOST type A"
    i_text_item_font.height = font_size
    i_text_item_font.ksu = 1
    i_text_item_array.ksAddArrayItem(-1, i_text_item_param)
    i_text_line_param.SetTextItemArr(i_text_item_array)

    iDocument2D.ksTextLine(i_text_line_param)
    iDocument2D.ksEndObj()


def m_to_mm(m: Tuple[float, ...] | List[float]) -> Tuple[float, ...] | Tuple[float, float]:
    return tuple(map(lambda i: i * 1000, m))


def mm_to_m(m: Tuple[float, ...]) -> Tuple[float, ...] | Tuple[float, float]:
    return tuple(map(lambda i: i / 1000, m))


def endpoint_by_distance_and_angle(starting_point: Union[Tuple[float, float], List[float]], distance: float,
                                   angle: float) -> Tuple[float, float]:
    xx = starting_point[0] + (distance * cos(np.radians(angle)))
    yy = starting_point[1] + (distance * sin(np.radians(angle)))
    return xx, yy


def angle_trunc(a: float) -> float:
    while a < 0.0:
        a += pi * 2
    return a


def get_angle_between_points(x_orig: float, y_orig: float, x_landmark: float, y_landmark: float) -> float:
    delta_y = y_landmark - y_orig
    delta_x = x_landmark - x_orig
    return angle_trunc(atan2(delta_y, delta_x))


def interpolate_line(first_point: str, second_point: str, interpolated_points: dict, point_dict: dict):
    _first_point = np.array(point_dict[str(first_point)])
    _second_point = np.array(point_dict[str(second_point)])

    _coordinate_dif = _first_point[:2] - _second_point[:2]

    _length = round(np.sqrt(np.sum(np.power(_coordinate_dif, 2))))

    _yia = [0, _length]
    _xia = [_first_point[2], _second_point[2]]

    _x = [0, _length]
    _f = interpolate.interp1d(_xia, _yia)

    _alpha = np.rad2deg(
        get_angle_between_points(_first_point[0], _first_point[1], _second_point[0], _second_point[1]))

    # if np.diff(sorted([int(round(_second_point[2])), int(ceil(_first_point[2]))])) <= 1:
    #     return

    _height_dif_sorted = sorted([_second_point[2], _first_point[2]])

    _height_dif = range(int(ceil(_height_dif_sorted[0])), int(_height_dif_sorted[1]) + 1)

    print(f"Points: {first_point, second_point}\n Range: {_height_dif}\n {_height_dif_sorted}\n")

    for i in _height_dif:
        _d = _f(i)
        _interpolated_point = endpoint_by_distance_and_angle(m_to_mm(_first_point[:2]), _d * 1000, _alpha)
        iDocument2D.ksPoint(*_interpolated_point, 0)
        add_text(f"{i}", _interpolated_point[0] - 10000, _interpolated_point[1] - 8000, color=14417715)

        if i in interpolated_points:
            interpolated_points[i].append(_interpolated_point)
        else:
            interpolated_points.update({i: [_interpolated_point]})
        print(f"{i}: {_interpolated_point}")
    print("\n")

    return interpolated_points


# https://stackoverflow.com/a/67313571
def get_intersections(x0, y0, r0, x1, y1, r1):
    # circle 1: (x0, y0), radius r0
    # circle 2: (x1, y1), radius r1

    d = sqrt((x1 - x0) ** 2 + (y1 - y0) ** 2)

    # non-intersecting
    if d > r0 + r1:
        return {}
    # One circle within other
    if d < abs(r0 - r1):
        return {}
    # coincident circles
    if d == 0 and r0 == r1:
        return {}
    else:
        a = (r0 ** 2 - r1 ** 2 + d ** 2) / (2 * d)
        h = sqrt(r0 ** 2 - a ** 2)
        x2 = x0 + a * (x1 - x0) / d
        y2 = y0 + a * (y1 - y0) / d
        x3 = x2 + h * (y1 - y0) / d
        y3 = y2 - h * (x1 - x0) / d
        x4 = x2 - h * (y1 - y0) / d
        y4 = y2 + h * (x1 - x0) / d
        return x3, y3, x4, y4


# %%
def add_rect(starting_corner: Tuple[float, float], h: float, w: float, angle: float = 0, style: int = 1) -> int:
    i_rectangle_param = kompas6_api5_module.ksRectangleParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_RectangleParam))
    i_rectangle_param.Init()
    i_rectangle_param.x = starting_corner[0]
    i_rectangle_param.y = starting_corner[1]
    i_rectangle_param.ang = angle
    i_rectangle_param.height = h
    i_rectangle_param.width = w
    i_rectangle_param.style = style
    return iDocument2D.ksRectangle(i_rectangle_param)


def _f_slope(x1, y1, x2, y2):  # Line slope given two points:
    return (y2 - y1) / (x2 - x1)


def _f_angle(s1, s2):
    return degrees(atan((s2 - s1) / (1 + (s2 * s1))))


def draw_meadow(point: Tuple[float, float], style: int = 2, size: float = 1000) -> Tuple[int, int]:
    line_one = iDocument2D.ksLineSeg(*point, *map(lambda i, j: (i + j), point, (0, 2.5 * size)), style)
    line_two = iDocument2D.ksLineSeg(*map(lambda i, j: (i + j), point, (1 * size, 0)),
                                     *map(lambda i, j: (i + j), point, (1 * size, 2.5 * size)), style)

    return line_one, line_two


def line_len(p1: Tuple[float, float] | List[float], p2: Tuple[float, float] | List[float]) -> float:
    return sum(map(lambda fp, sp: (sp - fp) ** 2, p1, p2)) ** 0.5


def sum_tuple(t1: Tuple[float, float] | List[float], t2: Tuple[float, float]) -> Tuple[float, float]:
    return t1[0] + t2[0], t1[1] + t2[1]


def add_raster(path: str, point: Tuple[float, float], scale: float = 1, angle: float = 0, embed: bool = True) -> None:
    """Добавляет растровое изображение в чертеж.

    :param path: Путь к файлу.
    :param point: Точка вставки (1 равняется кратности вида).
    :param scale: Масштаб вставки.
    :param angle: Угол поворота вставки.
    :param embed: Встраивать или нет.
    """
    iRasterParam = kompas6_api5_module.ksRasterParam(kompas_object.GetParamStruct(kompas6_constants.ko_RasterParam))
    iRasterParam.Init()
    iRasterParam.embeded = embed
    iRasterParam.fileName = str(Path(path).absolute())
    iPlacementParam = kompas6_api5_module.ksPlacementParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_PlacementParam))
    iPlacementParam.Init()
    iPlacementParam.angle = angle
    iPlacementParam.scale_ = scale
    iPlacementParam.xBase = point[0]
    iPlacementParam.yBase = point[1]
    iRasterParam.SetPlace(iPlacementParam)
    iDocument2D.ksInsertRaster(iRasterParam)


ONE_TO_SCALE = 2000


def main() -> None:
    for _variant in range(22, 23):
        wb = xw.Book('../../Геодезия.xlsm')
        ws = wb.sheets["Варианты"]

        for _letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            ws[f"{_letter}{3}"].value = ws[f"{_letter}{4 + _variant}"].value

        wb.save()
        wb.close()

        # %%
        table_79 = pd.read_excel('../../Геодезия.xlsm', sheet_name='pandasPlan79')
        table_10 = pd.read_excel('../../Геодезия.xlsm', sheet_name='pandasPlan10')

        add_view("План", 1 / ONE_TO_SCALE)
        add_layer(1, 3, "Основные Точки")

        for i, row in table_79.iterrows():
            # X и Y перевёрнуты
            iDocument2D.ksCircle(row["Y"] * 1000, row["X"] * 1000, 1000, 1)
            add_text(str(row["№№"]) + " (" + str("{:1.2f}".format(row["H"])).replace(".", ",") + ")",
                     row["Y"] * 1000 - 15000,
                     row["X"] * 1000 - 7000)

        add_layer(2, 3, "Побочные Точки")

        extra_points = {}

        for i, row in table_10.iterrows():
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
                extra_points.update({str(origin_label): _i_am_shit_at_coding})
                continue
            else:
                if not origin:
                    continue
            _d = row["Отсчет по дальномеру kl, м"]
            _alpha = row["Отсчет по гориз. Кругу"]
            xx = origin[0] + (_d * cos(np.radians(angle - _alpha)))
            yy = origin[1] + (_d * sin(np.radians(angle - _alpha)))
            # obj = iDocument2D.ksLineSeg(origin[0] * 1000, origin[1] * 1000, xx * 1000,yy * 1000, 6)

            add_text(
                str(row["№№"]) + " (" + str("{:1.3f}".format(row["Отметки реечн. точек Hр.т., м"])).replace(".",
                                                                                                            ",") + ")",
                xx * 1000 + 2000, yy * 1000 - 1000)

            iDocument2D.ksCircle(xx * 1000, yy * 1000, 1000, 1)
            extra_points.update({str(row["№№"]): [xx, yy, row["Отметки реечн. точек Hр.т., м"]]})

        interpolated_points = {}

        add_layer(3, 2, "Точки интерполяции пашни")
        # todo: переделать эту часть кода

        con_111 = "5,6,7,8,9,112,ПЗ41".split(",")

        center_point = np.array(extra_points["111"])
        for point in con_111:
            point = np.array(extra_points[point])

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

        # %%
        x = [extra_points[p][0] for p in extra_points]
        y = [extra_points[p][1] for p in extra_points]
        centroid = (sum(x) / len(extra_points), sum(y) / len(extra_points))
        iDocument2D.ksPoint(*(cord * 1000 for cord in centroid), 0)

        current_view = kompas_document_2d.ViewsAndLayersManager.Views.ActiveView
        current_view.X = 594 / 2 - centroid[0] / 2
        current_view.Y = 420 / 2 - centroid[1] / 2
        current_view.Update()

        # %%

        interpolate_line("7", "8", interpolated_points, extra_points)

        # %%
        add_layer(4, 3, "Кривые интерполяции пашни")

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
        # %%
        add_layer(5, 2, "Точки интерполяции берега")

        interpolated_points_shore = {}

        # ПЗ 41
        interpolate_line("4", "5", interpolated_points_shore, extra_points)
        interpolate_line("ПЗ41", "1", interpolated_points_shore, extra_points)
        interpolate_line("ПЗ41", "2", interpolated_points_shore, extra_points)
        interpolate_line("ПЗ41", "4", interpolated_points_shore, extra_points)
        interpolate_line("ПЗ41", "5", interpolated_points_shore, extra_points)
        interpolate_line("ПЗ41", "21", interpolated_points_shore, extra_points)

        # 112
        # interpolate_line("8", "10", interpolated_points_shore, extra_points)
        # interpolate_line("112", "10", interpolated_points_shore, extra_points)
        # interpolate_line("112", "13", interpolated_points_shore, extra_points)
        # interpolate_line("9", "12", interpolated_points_shore, extra_points)
        interpolate_line("12", "13", interpolated_points_shore, extra_points)
        # interpolate_line("13", "1", interpolated_points_shore, extra_points)
        interpolate_line("1", "112", interpolated_points_shore, extra_points)

        # 113
        interpolate_line("10", "14", interpolated_points_shore, extra_points)
        # interpolate_line("113", "15", interpolated_points_shore, extra_points)
        # interpolate_line("17", "113", interpolated_points_shore, extra_points)
        interpolate_line("16", "113", interpolated_points_shore, extra_points)
        # interpolate_line("17", "112", interpolated_points_shore, extra_points)
        # interpolate_line("17", "14", interpolated_points_shore, extra_points)

        # ПЗ 42
        interpolate_line("ПЗ42", "20", interpolated_points_shore, extra_points)
        interpolate_line("ПЗ42", "18", interpolated_points_shore, extra_points)
        interpolate_line("23", "22", interpolated_points_shore, extra_points)
        interpolate_line("22", "21", interpolated_points_shore, extra_points)
        interpolate_line("20", "21", interpolated_points_shore, extra_points)
        interpolate_line("1", "21", interpolated_points_shore, extra_points)
        interpolate_line("21", "21", interpolated_points_shore, extra_points)

        # %%
        add_layer(6, 3, "Кривые интерполяции берега")

        for height in interpolated_points_shore:

            if len(interpolated_points_shore[height]) < 3:
                continue

            interpolated_points_shore[height].sort()

            iDocument2D.ksBezier(False, 1)
            for point in interpolated_points_shore[height]:
                iDocument2D.ksPoint(*point, 0)
            iDocument2D.ksEndObj()

        # Горизонталь вокруг ФС
        add_layer(7, 2, "Точки интерполяции ФС")

        interpolated_points_fruit_garden = {}

        _interpolation_21_23 = interpolate_line("21", "23", {}, extra_points)
        _last_point_between_21_23 = list(_interpolation_21_23.keys())[-1]

        interpolated_points_fruit_garden.update(
            {_last_point_between_21_23: _interpolation_21_23[_last_point_between_21_23]})

        _interpolation_112_17 = interpolate_line("112", "17", {}, extra_points)
        if _last_point_between_21_23 in _interpolation_112_17:
            interpolated_points_fruit_garden[_last_point_between_21_23].append(
                *_interpolation_112_17[_last_point_between_21_23])

        _interpolation_14_17 = interpolate_line("14", "17", {}, extra_points)
        interpolated_points_fruit_garden[_last_point_between_21_23].append(
            *_interpolation_14_17[_last_point_between_21_23])

        _interpolation_15_113 = interpolate_line("15", "113", {}, extra_points)
        interpolated_points_fruit_garden[_last_point_between_21_23].append(
            *_interpolation_15_113[_last_point_between_21_23])

        _interpolation_p42_18 = interpolate_line("ПЗ42", "18", {}, extra_points)
        interpolated_points_fruit_garden[_last_point_between_21_23].insert(0, *_interpolation_p42_18[
            _last_point_between_21_23])

        # %%
        add_layer(8, 3, "Кривые интерполяции ФС")

        iDocument2D.ksBezier(False, 1)
        for point in interpolated_points_fruit_garden[_last_point_between_21_23]:
            iDocument2D.ksPoint(*point, 0)
        iDocument2D.ksEndObj()
        # %%
        # Горизонталь за ФС
        add_layer(9, 2, "Точки интерполяции за ФС")

        interpolated_points_behind_fruit_garden = {}

        _interpolation_112_10 = interpolate_line("112", "10", {}, extra_points)
        _last_point_between_112_10 = list(_interpolation_112_10.keys())[-1]

        interpolated_points_behind_fruit_garden.update(
            {_last_point_between_112_10: _interpolation_112_10[_last_point_between_112_10]})

        _interpolation_14_17 = interpolate_line("14", "17", {}, extra_points)
        interpolated_points_behind_fruit_garden[_last_point_between_112_10].append(
            *_interpolation_14_17[_last_point_between_112_10])

        _interpolation_15_113 = interpolate_line("15", "113", {}, extra_points)
        if _last_point_between_112_10 in _interpolation_15_113:
            interpolated_points_behind_fruit_garden[_last_point_between_112_10].append(
                *_interpolation_15_113[_last_point_between_112_10])

        _interpolation_8_10 = interpolate_line("8", "10", {}, extra_points)
        interpolated_points_behind_fruit_garden[_last_point_between_112_10].insert(0, *_interpolation_8_10[
            _last_point_between_112_10])

        add_layer(10, 3, "Кривые интерполяции за ФС")

        iDocument2D.ksBezier(False, 1)
        for point in interpolated_points_behind_fruit_garden[_last_point_between_112_10]:
            iDocument2D.ksPoint(*point, 0)
        iDocument2D.ksEndObj()

        # Соединить основные границы

        add_layer(10, 3, "Левая граница")

        _left_border = "4,5,6,7".split(sep=",")

        for i in range(len(_left_border) - 1):
            # print(f"{ _left_border[i]}: {extra_points[ _left_border[i]]}\t"
            #       f"{ _left_border[i+1]}: {extra_points[ _left_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*extra_points[_left_border[i]][:2], *extra_points[_left_border[i + 1]][:2])], 4)
        # %%
        add_layer(11, 3, "Верхняя граница")

        _top_border = "7,8,10,14,15".split(sep=",")

        for i in range(len(_top_border) - 1):
            # print(f"{ _top_border[i]}: {extra_points[ _top_border[i]]}\t"
            #       f"{ _top_border[i+1]}: {extra_points[ _top_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*extra_points[_top_border[i]][:2], *extra_points[_top_border[i + 1]][:2])], 1)
        # %%
        add_layer(12, 3, "Граница пашни")

        _p_border = "5,ПЗ41,9,112,8".split(sep=",")

        for i in range(len(_p_border) - 1):
            # print(f"{ _p_border[i]}: {extra_points[ _p_border[i]]}\t"
            #       f"{ _p_border[i+1]}: {extra_points[ _p_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*extra_points[_p_border[i]][:2], *extra_points[_p_border[i + 1]][:2])], 4)

        add_layer(13, 3, "Река")

        _river_border = "4,2,21,20".split(sep=",")

        for i in range(len(_river_border) - 1):
            # print(f"{ _river_border[i]}: {extra_points[ _river_border[i]]}\t"
            #       f"{ _river_border[i+1]}: {extra_points[ _river_border[i+1]]}")
            iDocument2D.ksLineSeg(
                *[i * 1000 for i in (*extra_points[_river_border[i]][:2], *extra_points[_river_border[i + 1]][:2])], 1)

        # вторая сторона реки

        _alpha = np.rad2deg(get_angle_between_points(*extra_points["4"][:2], *extra_points["3"][:2]))

        _second_shore = []
        _d = 32
        for i in range(len(_river_border)):
            _second_shore.append(endpoint_by_distance_and_angle(
                m_to_mm(extra_points[_river_border[i]][:2]), _d * 1000, _alpha))

        for i in range(len(_second_shore) - 1):
            if i != len(_second_shore) - 2:
                iDocument2D.ksLineSeg(*_second_shore[i], *_second_shore[i + 1], 1)
            else:
                iDocument2D.ksLineSeg(*_second_shore[i], *(p * 1000 for p in extra_points["19"][:2]), 1)

        # Добавить подпись

        _xx = (extra_points["2"][0] + extra_points["21"][0]) / 2 * 1000 + (_d / 2 * 1000 * cos(np.radians(_alpha)))
        _yy = (extra_points["2"][1] + extra_points["21"][1]) / 2 * 1000 + (_d / 2 * 1000 * sin(np.radians(_alpha)))

        add_text("р. Соть", _xx, _yy,
                 np.rad2deg(get_angle_between_points(*extra_points["2"][:2], *extra_points["21"][:2])), 5)
        # %%
        # Автодорога

        add_layer(14, 3, "Автодорога")

        _alpha = np.rad2deg(get_angle_between_points(*extra_points["6"][:2], *extra_points["7"][:2]))

        _autobahn = []
        _d = 10
        for i in range(len(_top_border)):
            xx = extra_points[_top_border[i]][0] * 1000 + (_d * 1000 * cos(np.radians(_alpha)))
            yy = extra_points[_top_border[i]][1] * 1000 + (_d * 1000 * sin(np.radians(_alpha)))
            _autobahn.append((xx, yy))

        for i in range(len(_autobahn) - 1):
            obj = iDocument2D.ksLineSeg(*_autobahn[i], *_autobahn[i + 1], 4)

        # %%

        # %%

        # %%
        # Текст пашни
        add_layer(15, 3, "Текст пашни")

        add_text("Пашня", *m_to_mm(sum_tuple(extra_points["111"][:2], (-20, 45))), 0, 7)
        # %%
        # Текст лес
        add_layer(16, 3, "Текст лес")

        add_text("Лес", *m_to_mm(sum_tuple(extra_points["6"][:2], (-60, 0))), 0, 7)

        # %%

        # %%

        # %%
        # Фруктовый сад
        add_layer(17, 3, "Колодец")

        _alpha_14_15 = np.rad2deg(get_angle_between_points(*extra_points["14"][:2], *extra_points["15"][:2]))
        _alpha_113_P42 = np.rad2deg(get_angle_between_points(*extra_points["113"][:2], *extra_points["ПЗ42"][:2]))

        # _well_angles = degrees(atan2(14.62,9.15))
        #
        _well_point_1 = m_to_mm(endpoint_by_distance_and_angle(extra_points["113"][:2], 40, _alpha_113_P42))
        # _obj_well_point_1 = iDocument2D.ksPoint(*_well_point_1, 0)

        _well_point_2 = m_to_mm(endpoint_by_distance_and_angle(extra_points["113"][:2], 58.61, _alpha_113_P42))
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

        add_layer(18, 3, "Фруктовый сад")

        # fruit_garden_points = [endpoint_by_distance_and_angle(extra_points["113"][:2], 81.5, _alpha_113_P42 - (44 +
        # (1/60))), endpoint_by_distance_and_angle(endpoint_by_distance_and_angle(extra_points["113"][:2], 58.61,
        # _alpha_113_P42), -8.2, _alpha_14_15), endpoint_by_distance_and_angle(endpoint_by_distance_and_angle(
        # extra_points["113"][:2], 133.41, _alpha_113_P42), -7.81, _alpha_14_15), endpoint_by_distance_and_angle(
        # extra_points["ПЗ42"][:2], 96.15, 0 - _alpha_113_P42 - (25 + (11/60)))]

        fruit_garden_points = [tuple(extra_points["17"][:2]), endpoint_by_distance_and_angle(
            endpoint_by_distance_and_angle(extra_points["113"][:2], 58.61, _alpha_113_P42), -8.2, _alpha_14_15),
                               endpoint_by_distance_and_angle(
                                   endpoint_by_distance_and_angle(extra_points["113"][:2], 133.41, _alpha_113_P42),
                                   -7.81,
                                   _alpha_14_15), tuple(extra_points["23"][:2])]

        iDocument2D.ksLineSeg(*m_to_mm(fruit_garden_points[0]), *m_to_mm(fruit_garden_points[3]), 1)
        for i in range(len(fruit_garden_points) - 1):
            iDocument2D.ksLineSeg(*m_to_mm(fruit_garden_points[i]), *m_to_mm(fruit_garden_points[i + 1]), 1)

        # _mid_fs_text = tuple(i/2 for i in (fruit_garden_points[0][0] + fruit_garden_points[1][0],
        # fruit_garden_points[0][1] + fruit_garden_points[1][1])) *map(lambda i, j: (i + j) * 1000, _mid_fs_text,
        # (-60, 0))

        _alpha_fs = np.rad2deg(get_angle_between_points(*fruit_garden_points[1], *fruit_garden_points[2]))

        add_text("ФС", *m_to_mm(endpoint_by_distance_and_angle(fruit_garden_points[0], 15, _alpha_fs + 30)),
                 np.rad2deg(get_angle_between_points(*fruit_garden_points[0], *fruit_garden_points[1])), 5)

        add_layer(19, 3, "2КЖ")

        _kg2_p1 = m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(extra_points["113"][:2], 80.05, _alpha_113_P42),
                -8.03, _alpha_14_15))
        _kg2_p2 = m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(extra_points["113"][:2], 110.23, _alpha_113_P42),
                -7.91, _alpha_14_15))

        add_rect(_kg2_p1, -16.05 * 1000, (110.23 - 80.05) * 1000, _alpha_fs)

        # *tuple(i/2 for i in (_kg2_p1[0] + _kg2_p2[0], _kg2_p1[1] + _kg2_p2[1]))

        add_text("2кж", _kg2_p2[0] - 2400, _kg2_p2[1] + 4000, np.rad2deg(get_angle_between_points(*_kg2_p2, *_kg2_p1)),
                 5)

        # %%
        # Железная дорога
        add_layer(19, 3, "Железная дорога")

        _railroad_p1 = endpoint_by_distance_and_angle(extra_points["113"][:2], 12.64, _alpha_14_15)
        _railroad_end = endpoint_by_distance_and_angle(
            endpoint_by_distance_and_angle(extra_points["113"][:2], 162.1, _alpha_113_P42), 14.28, _alpha_14_15)

        _alpha_rail = np.rad2deg(get_angle_between_points(*_railroad_p1, *_railroad_end))

        _railroad_p0 = endpoint_by_distance_and_angle(_railroad_p1, -70, _alpha_rail)

        iDocument2D.ksLineSeg(*m_to_mm(_railroad_p0), *m_to_mm(_railroad_end), 1)

        for i in range(5, int(line_len(_railroad_p0, _railroad_end)) + 5, 5):
            _start_point = endpoint_by_distance_and_angle(_railroad_p0, i, _alpha_rail)
            _side_point_1 = endpoint_by_distance_and_angle(_start_point, 2.5, 45 - _alpha_rail)
            _side_point_2 = endpoint_by_distance_and_angle(_start_point, -2.5, 45 - _alpha_rail)
            iDocument2D.ksLineSeg(*m_to_mm(_side_point_1), *m_to_mm(_side_point_2), 1)

        # %%
        # Условные обозначения леса, луга, фруктового сада.

        _plan_poly = [m_to_mm(tuple(extra_points["20"][:2])), m_to_mm(tuple(extra_points["21"][:2])),
                      m_to_mm(tuple(extra_points["2"][:2])),
                      m_to_mm(tuple(extra_points["4"][:2])), m_to_mm(tuple(extra_points["5"][:2])),
                      m_to_mm(tuple(extra_points["6"][:2])),
                      m_to_mm(tuple(extra_points["7"][:2])), m_to_mm(tuple(extra_points["8"][:2])),
                      m_to_mm(tuple(extra_points["10"][:2])),
                      m_to_mm(tuple(extra_points["14"][:2])), m_to_mm(tuple(extra_points["15"][:2])),
                      m_to_mm(tuple(extra_points["113"][:2])),
                      m_to_mm(tuple(extra_points["ПЗ42"][:2]))]

        _plan_poly_line = geometry.LineString(_plan_poly)
        _plan_poly = geometry.Polygon(_plan_poly_line)

        _farm_poly = [m_to_mm(tuple(extra_points["ПЗ41"][:2])), m_to_mm(tuple(extra_points["5"][:2])),
                      m_to_mm(tuple(extra_points["6"][:2])),
                      m_to_mm(tuple(extra_points["7"][:2])), m_to_mm(tuple(extra_points["8"][:2])),
                      m_to_mm(tuple(extra_points["112"][:2])),
                      m_to_mm(tuple(extra_points["9"][:2]))]

        _farm_poly_line = geometry.LineString(_farm_poly)
        _farm_poly = geometry.Polygon(_farm_poly_line)

        _fruit_garden_poly = geometry.Polygon(geometry.LineString((m_to_mm(i) for i in fruit_garden_points)))

        _fruit_garden_poly = affinity.scale(_fruit_garden_poly, 0.95, 0.95)

        _kg_poly = [_kg2_p1, m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(extra_points["113"][:2], 80.05, _alpha_113_P42),
                -8.03 - 16.05, _alpha_14_15)), m_to_mm(
            endpoint_by_distance_and_angle(
                endpoint_by_distance_and_angle(extra_points["113"][:2], 110.23, _alpha_113_P42),
                -7.91 - 16.05, _alpha_14_15)), _kg2_p2]

        _kg_poly_line = geometry.LineString(_kg_poly)
        _kg_poly = geometry.Polygon(_kg_poly_line)

        # %%
        add_layer(20, 0, "Условные обозначения луга")

        for i in range(int(extra_points["20"][1] - int(extra_points["20"][1]) % 5), int(extra_points["7"][1]), 20):
            for b in range(int(extra_points["4"][0] - int(extra_points["4"][1]) % 5), int(extra_points["15"][0]), 20):
                _d_point = geometry.Point(m_to_mm((b, i)))
                if _plan_poly.contains(_d_point) and not _farm_poly.contains(
                        _d_point) and not _fruit_garden_poly.contains(
                    _d_point):
                    draw_meadow(m_to_mm((b, i)))

        # %%
        add_layer(21, 0, "Условные обозначения фруктового сада")

        for i in range(int(fruit_garden_points[2][1]), int(fruit_garden_points[0][1]), 5):
            for b in range(int(fruit_garden_points[-1][0]), int(fruit_garden_points[1][0]), 5):
                _d_point = geometry.Point(m_to_mm((b, i)))
                if _fruit_garden_poly.contains(_d_point) and not _kg_poly.contains(_d_point):
                    iDocument2D.ksCircle(*m_to_mm((b, i)), 1000, 2)
        # %%
        add_layer(22, 0, "Условные обозначения леса")

        for i in range(int(extra_points["4"][1] - int(extra_points["4"][1]) % 5), int(extra_points["7"][1]), 20):
            for b in range(int(extra_points["3"][0] - int(extra_points["3"][1]) % 5), int(extra_points["7"][0]), 20):
                _d_point = geometry.Point(m_to_mm((b, i)))
                if not _plan_poly.contains(_d_point):
                    iDocument2D.ksCircle(*m_to_mm((b, i)), 2500, 2)

        # %%

        # %%
        # Дерево
        add_layer(23, 0, "Дерево")

        _a1 = radians(49 + 15 / 60)
        _a2 = radians(_f_angle(_f_slope(*extra_points["113"][:2], *extra_points["112"][:2]),
                               _f_slope(*extra_points["ПЗ42"][:2], *extra_points["113"][:2])) - (55 + 6 / 60))

        s = line_len(extra_points["113"][:2], extra_points["112"][:2])

        _a3 = pi - _a2 - _a1

        _l1 = (s / sin(_a3)) * sin(_a1)
        _l2 = (s / sin(_a3)) * sin(_a2)

        _tree = endpoint_by_distance_and_angle(extra_points["113"][:2], _l1, degrees(_a2) + np.rad2deg(
            get_angle_between_points(*extra_points["113"][:2], *extra_points["112"][:2])))

        _tree_point = iDocument2D.ksPoint(*m_to_mm(_tree), 0)

        add_raster('tree-64x64.png', (_tree[0] / 2 - 1.847273, _tree[1] / 2), 0.025)


        # %%
        # Рамка
        add_layer(24, 0, "Рамка")

        _padding = 30
        _top_left_frame_corner = m_to_mm((extra_points["3"][0] - _padding, extra_points["7"][1] + _padding))
        _bottom_right_frame_corner = m_to_mm((extra_points["15"][0] + _padding, extra_points["19"][1] - _padding))

        _wh_dif = (-1 * (_top_left_frame_corner[1] - _bottom_right_frame_corner[1]),
                   -1 * (_top_left_frame_corner[0] - _bottom_right_frame_corner[0]))
        _max_dim = max((abs(i) for i in _wh_dif))
        _wh_dif = (-1 * _max_dim, _max_dim)

        # iDocument2D.ksPoint(*_top_left_frame_corner, 0)
        # iDocument2D.ksPoint(*_bottom_right_frame_corner, 0)

        add_rect(_top_left_frame_corner, *_wh_dif)

        _padding += _padding
        _top_left_frame_corner = m_to_mm((extra_points["3"][0] - _padding, extra_points["7"][1] + _padding))
        _bottom_right_frame_corner = m_to_mm((extra_points["15"][0] + _padding, extra_points["19"][1] - _padding))

        _wh_dif = (-1 * (_top_left_frame_corner[1] - _bottom_right_frame_corner[1]),
                   -1 * (_top_left_frame_corner[0] - _bottom_right_frame_corner[0]))
        _max_dim = max((abs(i) for i in _wh_dif))
        _wh_dif = (-1 * _max_dim, _max_dim)

        add_rect(_top_left_frame_corner, *_wh_dif)

        # %%
        _main_points = [tuple(int(i // 100) / 10 for i in extra_points[point][:2]) for point in
                        "ПЗ41,111,112,113,ПЗ42".split(sep=",")]

        _min_x = min(_main_points, key=lambda x: x[0])[0]
        _min_y = min(_main_points, key=lambda y: y[1])[1]

        _x_list = [round(_min_x + 0.2 * i, 1) for i in range(3)]
        _y_list = [round(_min_y + 0.2 * i, 1) for i in range(3)]

        for x_cord in _x_list:
            for y_cord in _y_list:
                iDocument2D.ksPoint(*m_to_mm(m_to_mm((x_cord, y_cord))), 8)

        # %%
        for cord in _x_list:
            add_text(str(cord), int(cord * 1000000) + 2.5 * 2000, _bottom_right_frame_corner[1] + _padding / 7 * 1000,
                     90, 5)
            add_text(str(cord), int(cord * 1000000) + 2.5 * 2000, _top_left_frame_corner[1] - _padding / 2.5 * 1000, 90,
                     5)

        # %%
        for cord in _y_list:
            add_text(str(cord), _top_left_frame_corner[0] + _padding / 7 * 1000, int(cord * 1000000) - 2.5 * 2000, 0, 5)
            add_text(str(cord), _bottom_right_frame_corner[0] + _padding / 2 * 1000, int(cord * 1000000) - 2.5 * 2000,
                     0, 5)

        # %%
        # Добавить подпись

        add_text(f"План Масштаб 1:{ONE_TO_SCALE} Вариант: {_variant}", _top_left_frame_corner[0] + 10000,
                 _bottom_right_frame_corner[1] - 40000, 0, 14)
        # %%
        if Path('../../../Watermark.png').exists():
            # Добавить картинку

            iRasterParam = kompas6_api5_module.ksRasterParam(
                kompas_object.GetParamStruct(kompas6_constants.ko_RasterParam))
            iRasterParam.Init()
            iRasterParam.embeded = True
            iRasterParam.fileName = str(Path('../../../Watermark.png').absolute())
            iPlacementParam = kompas6_api5_module.ksPlacementParam(
                kompas_object.GetParamStruct(kompas6_constants.ko_PlacementParam))
            iPlacementParam.Init()
            iPlacementParam.angle = 0
            iPlacementParam.scale_ = 0.4
            iPlacementParam.xBase = extra_points["3"][0] / 2
            iPlacementParam.yBase = extra_points["20"][1] / 2
            iRasterParam.SetPlace(iPlacementParam)
            iDocument2D.ksInsertRaster(iRasterParam)
        # %%
        kompas_document.SaveAs(str(Path(f'../{_variant}.pdf').absolute()))


if __name__ == "__main__":
    main()