# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
if sys.version_info.major == 2:
    from cStringIO import StringIO as ImageIO
else:
    from io import BytesIO as ImageIO
    from functools import reduce
from PIL import Image
import requests
import time


def avhash(im):
    if not isinstance(im, Image.Image):
        im = Image.open(im)
    im = im.resize((49, 49), Image.ANTIALIAS).convert('L')
    avg = reduce(lambda x, y: x + y, im.getdata()) / 2401.
    return reduce(lambda x, (y, z): x | (z << y),
                  enumerate(map(lambda i: 0 if i < avg else 1, im.getdata())),
                  0)


def hamming(h1, h2):
    h, d = 0, h1 ^ h2
    while d:
        h += 1
        d &= d - 1
    return h

UserAgent = 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36'

Pwd = os.path.dirname(__file__).decode('gbk')

fonthash = [avhash(os.path.join(Pwd, 'font', '{0}.png'.format(i))) for i in range(1, 10)]

driver = 'chrome'

if driver == 'phantomjs':
    capabilities = DesiredCapabilities.PHANTOMJS.copy()
    capabilities['phantomjs.page.settings.userAgent'] = UserAgent
    capabilities['phantomjs.page.settings.loadImages'] = True
    browser = webdriver.PhantomJS(executable_path=os.path.join(os.environ['SELENIUM'], 'phantomjs.exe' if sys.platform == 'win32' else 'phantomjs'), desired_capabilities=capabilities)
elif driver == 'chrome' :
    browser = webdriver.Chrome()

browser.get('http://fuzhou.baixing.com/oz/s9verify')

print(browser.find_element_by_css_selector('h3').text.strip())

verify = browser.find_element_by_css_selector('img#ez-verify-image')

s9verify = Image.open(ImageIO(requests.get(verify.get_attribute('src'), headers={'User-Agent': UserAgent}).content))
s9verify.save(os.path.join(Pwd, '{0}.png'.format(int(time.time()))))

'''图片九宫格坐标'''
coor = [
    (0,0,50,50), (51,0,100,50), (101,0,150,50), 
    (0,51,50,100), (51,51,100,100), (101,51,150,100), 
    (0,101,50,150), (51,101,100,150), (101,101,150,150)
]

OCR = [[], []]
for c in coor:
    imghash = avhash(s9verify.crop(c))
    result = map(lambda x: hamming(x, imghash), fonthash)
    OCR[0].append(result.index(min(result)))
    OCR[1].append(min(result))

DifferenceSet = list(set(range(9)) - set(OCR[0]))

if len(DifferenceSet) == 1:
    x, y = [m for m,n in enumerate([OCR[0].count(i) for i in OCR[0]]) if n == 2]
    if OCR[1][x] < OCR[1][y]:
        OCR[0][y] = DifferenceSet[0]
    else:
        OCR[0][x] = DifferenceSet[0]


print('九宫格识别')
for line in [map(lambda x: x + 1, OCR[0][ 3 * i: 3 * i +3]) for i in range(3)]:
    print(line)

'''生成九宫格点击坐标'''
Points = [(x, y) for y in [25, 75, 125] for x in [25, 75, 125]]

Codes = map(lambda x: int(x) - 1, filter(lambda x: x.isdigit(), browser.find_element_by_css_selector('h5#ez-verify-title > i').text.split()))

print('需要依次点击: [{0}]'.format(' - '.join(map(lambda x: str(x + 1), Codes))))

for code in Codes:
    x, y = Points[OCR[0].index(code)]
    ActionChains(browser).move_to_element_with_offset(verify, 0, 0).move_by_offset(x, y).click().perform()
    time.sleep(1)
print(browser.title)
time.sleep(5)
browser.quit()
del browser
