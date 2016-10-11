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


testData = {
    'title': u'X公司 - 每月销售额饼形图', 
    'xaxisname': u'月份',
    'data': {
        'xaxis': [u'一月', u'二月', u'三月', u'四月', u'五月', u'六月', u'七月', u'八月', u'九月', u'十月', u'十一月', u'十二月'], 
        'values': [462, 300, 671, 494, 761, 960, 629, 622, 376, 494, 761, 650]
    },
    'length': 12
}


Pwd = os.path.dirname(__file__).decode('gbk')

root = lxml.etree.Element('graph')
root.set('caption', testData['title'])
root.set('xAxisName', testData['xaxisname'])
root.set('showNames', '1')
root.set('decimalPrecision', '0')
root.set('formatNumberScale', '0')

for i in range(testData['length']):
    child = lxml.etree.SubElement(root, 'set')
    child.set('name', testData['data']['xaxis'][i])
    child.set('value', str(testData['data']['values'][i]))
    '''设置随机颜色代码'''
    child.set('color', ''.join(map(lambda x: '%02X' % x, [randint(0, 255) for j in range(3)])))

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

    OLEObject = Sht.Shapes.AddOLEObject(ClassType='ShockwaveFlash.ShockwaveFlash', Link=False, DisplayAsIcon=False, Left=10, Top=10, Width=400, Height=350).OLEFormat.Object.Object
    OLEObject.Base = os.path.join(Pwd, 'Pie2D.swf')
    OLEObject.Movie = os.path.join(Pwd, 'Pie2D.swf')
    OLEObject.EmbedMovie = True
    OLEObject.FlashVars = 'dataXML={0}'.format(lxml.etree.tounicode(root).encode('gbk'))

    #print lxml.etree.tostring(root, pretty_print=True, encoding='utf-8').decode('utf-8')
    print lxml.etree.tounicode(root)

    tree = lxml.etree.ElementTree(root)
    tree.write(os.path.join(Pwd, 'Chart.xml'), pretty_print=True, xml_declaration=True, encoding='utf-8')

    xlWbk.SaveAs(os.path.join(Pwd, 'Chart.xlsx'), win32com.client.constants.xlOpenXMLWorkbook)
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
