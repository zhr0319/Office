# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
import os
import sys
import clr
import System
import System.Windows.Forms
import System.Drawing
import time

notify = System.Windows.Forms.NotifyIcon()
notify.Icon = System.Drawing.Icon(os.path.join(os.path.dirname(sys.executable), 'DLLs', 'py.ico'))
notify.Visible = True
print('别看我，看右下角')
notify.ShowBalloonTip(5000, '注意', '准备吃饭！Go! Go! Go!', System.Windows.Forms.ToolTipIcon.Warning);
time.sleep(10)
notify.Dispose()
del notify
