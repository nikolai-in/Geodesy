# -*- coding: utf-8 -*-
#|Высоты

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
MH.iKompasObject  = kompas_object

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))
MH.iApplication  = application


Documents = application.Documents
#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

hights = {
    0: [84.814180186008, 105.671],
    1: [94.814180186008, 103.133],
    2: [104.814180186008, 98.344],
    35: [108.314180186008, 96.278],
    3: [114.814180186008, 96.399],
    4: [124.814180186008, 101.017],
    71: [131.914180186008, 102.666],
    5: [134.814180186008, 101.164],
    6: [144.814180186008, 96.664],
    7: [154.814180186008, 94.869],
    8: [164.814180186008, 93.189],
    29: [167.714180186008, 91.182],
    9: [174.814180186008, 91.904],
    10: [184.814180186008, 94.100],
    75: [192.314180186008, 95.236],
    11: [194.814180186008, 93.974],
    12: [204.814180186008, 89.416],
    13: [214.814180186008, 87.174],
    14: [224.814180186008, 84.805],
    50: [229.814180186008, 83.203],
    15: [234.814180186008, 83.197],
    49: [239.714180186008, 83.991],
    16: [244.814180186008, 85.842],
    17: [254.814180186008, 86.990],
    18: [264.814180186008, 87.158],
    19: [274.814180186008, 89.786],
    72: [282.014180186008, 87.746],
    20: [284.814180186008, 89.914],
}

vertical = [137.076729641143 + i for i in range(0, 160, 10)]

print(horizontal,vertical,hights)