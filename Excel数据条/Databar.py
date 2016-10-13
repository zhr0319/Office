# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import lxml.etree
from random import randint
import win32com.client
import win32process
import psutil
import time
import gc


Pwd = os.path.dirname(__file__).decode('gbk')


try:
    xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
    Pid = win32process.GetWindowThreadProcessId(xlApp.Hwnd)[1]
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.DisplayStatusBar = False
    xlApp.EnableEvents = False
    xlApp.Interactive = False

    xlWbk = xlApp.Workbooks.Add()
    Sht = xlWbk.ActiveSheet

    Data = Sht.Range(Sht.Cells(1, 1), Sht.Cells(1, 10))
    Data.Value = [i for i in xrange(1, 11)]
    Databar = Data.FormatConditions.AddDatabar()
    Databar.MinPoint.Modify(newtype=win32com.client.constants.xlConditionValueAutomaticMin)
    Databar.MaxPoint.Modify(newtype=win32com.client.constants.xlConditionValueAutomaticMax)
    Databar.BarColor.Color = 49407
    Databar.BarBorder.Color.Color = 49407
    del Data, Databar

    for Sht in xlWbk.Worksheets:
        CellsCount = Sht.UsedRange.Cells.CountLarge if float(xlApp.Version) >= 15 else Sht.UsedRange.Cells.Count
        if CellsCount == 1: Sht.Delete()

    xlWbk.SaveAs(os.path.join(Pwd, 'Databar.xlsx'), win32com.client.constants.xlOpenXMLWorkbook)
except Exception as error:
    print error
finally:
    xlWbk.Close(SaveChanges=False)
    del Sht, xlWbk

    xlApp.Quit()
    del xlApp
    time.sleep(0.5)

    if psutil.pid_exists(Pid):
        os.kill(Pid, -9)
        print 'kill Excel'
    gc.collect()
