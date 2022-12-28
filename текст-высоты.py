# -*- coding: utf-8 -*-
# |Точки2

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH

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

# obj = iDocument2D.ksPoint(79.00655451613, 147.076729641143, 0)
# obj = iDocument2D.ksPoint(79.00655451613, 157.076729641143, 0)
# obj = iDocument2D.ksPoint(79.00655451613, 167.076729641143, 0)

x = 79
y = 146
n = 84

for i in range(1, 17):
    iParagraphParam = kompas6_api5_module.ksParagraphParam(
        kompas_object.GetParamStruct(kompas6_constants.ko_ParagraphParam))
    iParagraphParam.Init()
    iParagraphParam.x = x
    iParagraphParam.y = y
    iParagraphParam.ang = 0
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
    iTextItemParam.s = str(n)
    iTextItemParam.type = 0
    iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
    iTextItemFont.Init()
    iTextItemFont.bitVector = 4096
    iTextItemFont.color = 0
    iTextItemFont.fontName = "GOST type A"
    iTextItemFont.height = 2.5
    iTextItemFont.ksu = 1
    iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
    iTextLineParam.SetTextItemArr(iTextItemArray)

    iDocument2D.ksTextLine(iTextLineParam)
    obj = iDocument2D.ksEndObj()

    n += 2
    y += 10
