from math import *
from pathlib import Path
from typing import List, Tuple
from typing import Union

import LDefin2D  # Компас-3д
import MiscellaneousHelpers as miscHelpers  # Компас-3д
import numpy as np
import pythoncom  # Компас-3д
from scipy import interpolate
from win32com.client import Dispatch, gencache  # Компас-3д

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


def f_slope(x1, y1, x2, y2):  # Line slope given two points:
    return (y2 - y1) / (x2 - x1)


def f_angle(s1, s2):
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
    i_raster_param = kompas6_api5_module.ksRasterParam(kompas_object.GetParamStruct(kompas6_constants.ko_RasterParam))
    i_raster_param.Init()
    i_raster_param.embeded = embed
    i_raster_param.fileName = str(Path(path).absolute())
    iPlacementParam = kompas6_api5_module.ksPlacementParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_PlacementParam))
    iPlacementParam.Init()
    iPlacementParam.angle = angle
    iPlacementParam.scale_ = scale
    iPlacementParam.xBase = point[0]
    iPlacementParam.yBase = point[1]
    i_raster_param.SetPlace(iPlacementParam)
    iDocument2D.ksInsertRaster(i_raster_param)
