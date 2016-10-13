# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import cv2
import os
import win32gui, win32con


Pwd = os.path.dirname(__file__).decode('gbk')


def Main():
    try:
        photo = win32gui.GetOpenFileNameW(hwndOwner=None, hInstance=None, Filter=u'图片文件\0*.jpg;*.jpeg;*.png*\0', CustomFilter=None, FilterIndex=0, File=None, MaxFile=10240, InitialDir=Pwd, Title=u'选择需要处理图片', Flags=win32con.OFN_EXPLORER, DefExt=None, TemplateName=None)[0]
    except:
        return

    print u'读取原图'
    im = cv2.imread(photo.encode('gbk'))
    h, w = im.shape[:2]
    zoom = (w if w >= h else h) / 800.
    if zoom > 1:
        print u'调整图片尺寸'
        newimg = cv2.resize(im, (int(w * 1/zoom),int(h * 1/zoom)), interpolation=cv2.INTER_AREA)
    else:
        newimg = im

    print u'显示原图'
    cv2.imshow(u'人脸识别'.encode('gbk'), newimg);cv2.waitKey(5)

    print u'载入面部特征数据'
    classfier = cv2.CascadeClassifier(os.path.join(Pwd.encode('gbk'), 'face', 'haarcascade_frontalface_alt_tree.xml'))
    print u'开始标示'
    for face in classfier.detectMultiScale(cv2.cvtColor(newimg, cv2.COLOR_BGR2GRAY), 1.1, 2, cv2.CASCADE_SCALE_IMAGE, (20, 20)):
        x, y, w, h = face
        cv2.rectangle(newimg, (x, y), (x+w, y+h), (0, 255, 0))

    cv2.imshow(u'人脸识别'.encode('gbk'), newimg);cv2.waitKey(0)


if __name__ == '__main__':
    Main()
