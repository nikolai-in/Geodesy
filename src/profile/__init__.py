"""Продольный профиль на трассе"""
import pandas as pd
import math
#%%
# Импорт компосовской херни
import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH
#%%
#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                             pythoncom.IID_IDispatch))
MH.iKompasObject = kompas_object

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                             pythoncom.IID_IDispatch))
MH.iApplication = application

Documents = application.Documents
#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()
#%%
df = pd.read_excel('Геодезия.xlsm', sheet_name='pandas')

df
#%%
HORIZONTAL_TEXT_AXIS = [i - (84.814180186008 - 84.421252) for i in [84.814180186008, 94.814180186008, 104.814180186008, 108.314180186008, 114.814180186008,
                        124.814180186008, 131.914180186008, 134.814180186008, 144.814180186008, 154.814180186008,
                        164.814180186008, 167.714180186008, 174.814180186008, 184.814180186008, 192.314180186008,
                        194.814180186008, 204.814180186008, 214.814180186008, 224.814180186008, 229.814180186008,
                        234.814180186008, 239.714180186008, 244.814180186008, 254.814180186008, 264.814180186008,
                        274.814180186008, 282.014180186008, 284.814180186008,] ]

POINT_NAMES = [0,1,2,35,3,4,71,5,6,7,8,29,9,10,75,11,12,13,14,50,15,49,16,17,18,19,72,20,]

BASELINE = 137.703368

Y_MIN = math.floor(min(df["Отметки H, м"])) - 2
#%%
hights = [[HORIZONTAL_TEXT_AXIS[i], df["Отметки H, м"][i], df["ПО"][i]] for i in range(len(POINT_NAMES))]
#%%
def add_layer(id: int, active: bool = True, name: str = None):
    obj = iDocument2D.ksLayer(id)
    iLayerParam = kompas6_api5_module.ksLayerParam(kompas_object.GetParamStruct(kompas6_constants.ko_LayerParam))
    iLayerParam.Init()
    if name:
        iLayerParam.name = name
    iLayerParam.state = int(active)
    iDocument2D.ksSetObjParam(obj, iLayerParam, LDefin2D.ALLPARAM)


def add_text(text: str, x: float, y: float, angle: float = 0, font_size: float = 2.5):
    iParagraphParam = kompas6_api5_module.ksParagraphParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_ParagraphParam))
    iParagraphParam.Init()
    iParagraphParam.x = x
    iParagraphParam.y = y
    iParagraphParam.ang = angle
    iParagraphParam.height = 3.55
    iParagraphParam.width = 4
    iParagraphParam.hFormat = 0
    iParagraphParam.vFormat = 0
    iParagraphParam.style = 1
    iDocument2D.ksParagraph(iParagraphParam)

    iTextLineParam = kompas6_api5_module.ksTextLineParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
    iTextLineParam.Init()
    iTextLineParam.style = 1
    iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
    iTextItemParam = kompas6_api5_module.ksTextItemParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
    iTextItemParam.Init()
    iTextItemParam.iSNumb = 0
    iTextItemParam.s = text
    iTextItemParam.type = 0
    iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
    iTextItemFont.Init()
    iTextItemFont.bitVector = 4096
    iTextItemFont.color = 0
    iTextItemFont.fontName = "GOST type A"
    iTextItemFont.height = font_size
    iTextItemFont.ksu = 1
    iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
    iTextLineParam.SetTextItemArr(iTextItemArray)

    iDocument2D.ksTextLine(iTextLineParam)
    obj = iDocument2D.ksEndObj()


#%%
if input("Добавить вертикальную текстовую ось (y/n)? \n") == "y":
    add_layer(1100, True, "Текст Ось Y")
    x = 79
    y = 146.5
    n = Y_MIN + 2
    for i in range(1, 17):
        add_text(n, x, y)
        n += 2
        y += 10
#%%
if input("Добавить 'Отметки земли, м' (y/n)? \n") == "y":
    add_layer(1110, True, "Отметки земли, м")
    for point in hights:
        add_text("{:1.3f}".format(point[1]), point[0] + 1.2, 70.2, 90)
#%%
if input("Добавить 'Проектные отметки, м' (y/n)? \n") == "y":
    add_layer(1110, True, "Проектные отметки, м")
    for point in hights:
        add_text("{:1.3f}".format(point[2]), point[0] + 1.2, 95.2, 90)
#%%
if input("Рисовать вертикали? (y/n) \n") == "y":
    add_layer(1120, True, "Вертикали")
    for point in hights:
        obj = iDocument2D.ksLineSeg(point[0], BASELINE, point[0],BASELINE + (point[1] - Y_MIN)*5, 2)
    if input("Соединять вертикали? (y/n) \n") == "y":
        for i in range(len(hights)):
                if i == 0:
                    continue
                obj = iDocument2D.ksLineSeg(hights[i-1][0], BASELINE + (hights[i-1][1] - Y_MIN)*5, hights[i][0],BASELINE + (hights[i][1] - Y_MIN)*5, 2)

#%%
print('Done!')
#%%
if input("Добавить 'Отметки земли, м' (y/n)? \n") == "y":
    add_layer(1110, True, "Отметки земли, м")
    for point in hights:
        add_text("{:1.3f}".format(point[1]), point[0] + 1.2, 70.2, 90)