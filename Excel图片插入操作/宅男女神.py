# -*- coding: utf-8 -*-
from __future__ import unicode_literals, division
import win32com.client
import win32process
import psutil
import os
import gc
import numpy as np
import requests
import grequests
from bs4 import BeautifulSoup
from urlparse import urlparse
from PIL import Image
from random import choice, shuffle
import hashlib
import traceback
import time


Pwd = os.path.dirname(__file__).decode('gbk')


class constants:
    msoTrue = -1
    msoPictureCompressDocDefault = -1
    msoPictureCompressFalse = 0
    msoPictureCompressTrue = 1
    msoCTrue = 1
    msoFalse = 0
    msoTriStateMixed = -2
    msoTriStateToggle = -3


win32com.client.constants.__dicts__.append(constants.__dict__)


UserAgent = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36'

RowHeight = [40, 60, 80]
MergeCells = [True, False]

req = requests.Session()


xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
Pid = win32process.GetWindowThreadProcessId(xlApp.Hwnd)[1]

try:
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.DisplayStatusBar = False
    xlApp.EnableEvents = False
    xlApp.Interactive = False

    xlWbk = xlApp.Workbooks.Add()
    Sht = xlWbk.ActiveSheet
    xlApp.Calculation = win32com.client.constants.xlCalculationManual

    print '获取网站首页图片链接'
    rep = req.get('http://www.zngirls.com/', headers={'User-Agent': UserAgent})

    soup = BeautifulSoup(rep.content, 'lxml')
    urls = map(lambda x: x.attrs['data-original'].replace('_s.jpg', '.jpg'), soup.select('div.post_entry > ul > li.girlli img'))

    print '下载图片'
    Pics = []
    for x in grequests.map((grequests.get(url, headers={'User-Agent': UserAgent}) for url in urls)):
        if x.status_code == 200:
            PicPath = os.path.join(Pwd, 'Pics', urlparse(x.url).path[1:].replace('/', '\\'))
            Pics.append([PicPath])
            if not os.path.exists(PicPath):
                if not os.path.exists(os.path.dirname(PicPath)):
                    os.makedirs(os.path.dirname(PicPath), 0x777)
                open(PicPath, 'wb').write(x.content)

    print '复制数组内容并乱序处理'
    Pics = Pics * 2
    shuffle(Pics)
    ImageRange = Sht.Range('A1:A{0}'.format(len(Pics)))
    ImageRange.Value = Pics

    PicHash = {}

    print '插入图片到Excel中，并随机设置单元格的行高、随机合并单元格'
    for C in ImageRange:
        if C.Value is not None:
            C.RowHeight = choice(RowHeight)
            if choice(MergeCells):
                C.GetResize(ColumnSize=2).Merge()
            try:
                hashcode = hashlib.sha1(C.Value.lower().encode('utf-8')).hexdigest()
                CellsSize = np.asarray([C.MergeArea.Width, C.MergeArea.Height])

                if PicHash.has_key(hashcode):
                    pic = Sht.Shapes(PicHash[hashcode]).Duplicate()
                    ImageSize = np.asarray([pic.Width, pic.Height])

                    ratio = round((CellsSize / ImageSize).min() * 0.9, 4)

                    Width, Height = ImageSize * ratio

                    Left = C.Left + ((C.MergeArea.Width - Width) / 2)
                    Top = C.Top + ((C.MergeArea.Height - Height) / 2)

                    pic.Top, pic.Left, pic.Width, pic.Height = Top, Left, Width, Height
                else:
                    PicPath = C.Value
                    ImageSize = np.asarray(Image.open(PicPath).size)

                    ratio = round((CellsSize / ImageSize).min() * 0.9, 4)

                    Width, Height = ImageSize * ratio
                    Left = C.Left + ((C.MergeArea.Width - Width) / 2)
                    Top = C.Top + ((C.MergeArea.Height - Height) / 2)
                    name = os.path.basename(PicPath)

                    pic = Sht.Shapes.AddPicture2(PicPath.replace('/', '\\'), win32com.client.constants.msoFalse, win32com.client.constants.msoTrue, Left, Top, Width, Height, win32com.client.constants.msoPictureCompressTrue)
                    PicHash[hashcode] = name

                pic.Name = PicHash[hashcode]
                pic.Placement = win32com.client.constants.xlMoveAndSize
                pic.LockAspectRatio = win32com.client.constants.msoTrue
            except Exception as error:
                print traceback.format_exc()

    try:
        del C, pic, PicHash, Width, Height, ratio, ImageRange
    except:
        pass

    print '保存文件'
    xlApp.Calculation = win32com.client.constants.xlCalculationAutomatic
    xlWbk.SaveAs(os.path.join(Pwd, '宅男女神.xlsx'), win32com.client.constants.xlOpenXMLWorkbook)
except Exception as error:
    xlWbk.Close(SaveChanges=False)
finally:
    del Sht, xlWbk
    xlApp.Quit()
    del xlApp
    time.sleep(0.5)

    if psutil.pid_exists(Pid):
        os.kill(Pid, -9)
    gc.collect()
    time.sleep(5)
