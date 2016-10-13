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

    Data = Sht.Range('A1:A11')
    Data.Value = [['Title'], ] + [[i, ] for i in xrange(1, 11)]

    listobject = Sht.ListObjects.Add(SourceType=win32com.client.constants.xlSrcRange, Source=Data, XlListObjectHasHeaders=win32com.client.constants.xlYes, TableStyleName='TableStyleLight1')
    del Data

    listobject.Name = 'List_1'

    Sht.ListObjects('List_1')

    listobject.Range.AutoFilter(Field=1, Criteria1=[str(i) for i in xrange(1, 11) if i % 2 == 0], Operator=win32com.client.constants.xlFilterValues)

    xlApp.ActiveWorkbook.SlicerCaches.Add2(Source=listobject, SourceField='Title').Slicers.Add(SlicerDestination=Sht, Name='Slicer_1', Caption='Slicer_1', Top=Sht.Range('C1').Top, Left=Sht.Range('C1').Left, Width=150, Height=185)

    del listobject

    for Sht in xlWbk.Worksheets:
        CellsCount = Sht.UsedRange.Cells.CountLarge if float(xlApp.Version) >= 15 else Sht.UsedRange.Cells.Count
        if CellsCount == 1: Sht.Delete()

    xlWbk.SaveAs(os.path.join(Pwd, 'Slicer.xlsx'), win32com.client.constants.xlOpenXMLWorkbook)
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
