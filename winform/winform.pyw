# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import clr
import System
from System.Windows.Forms import Application, Form, Button, Label
from System.Windows.Forms import FormStartPosition, FormBorderStyle
from System.Windows.Forms import DateTimePicker, DateTimePickerFormat
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon
from System.Drawing import Size, Point, ContentAlignment
from System import DateTime


def dpicker1_ValueChanged(sender, event):
    if DateTime.Compare(dpicker1.Value, dpicker2.Value) == 1:
        dpicker2.Value = dpicker1.Value


def dpicker2_ValueChanged(sender, event):
    if DateTime.Compare(dpicker1.Value, dpicker2.Value) == 1:
        dpicker1.Value = dpicker2.Value


form = Form()
form.StartPosition = FormStartPosition.CenterScreen
form.Size = Size(400, 200)
form.FormBorderStyle = FormBorderStyle.Fixed3D
form.Text = '测试窗口'

label1 = Label()
label1.Text = '起始日期'
label1.Location = Point(10, 40)
label1.Width = 75
label1.TextAlign = ContentAlignment.MiddleCenter
form.Controls.Add(label1)

dpicker1 = DateTimePicker()
dpicker1.Tag = 'date1'
dpicker1.Location = Point(85, 40)
dpicker1.Size = Size(90, 50)
dpicker1.Format = DateTimePickerFormat.Custom
dpicker1.CustomFormat = 'yyyy-MM-dd'
dpicker1.Value = DateTime.Now.AddDays(-1)
dpicker1.ValueChanged += dpicker1_ValueChanged
form.Controls.Add(dpicker1)


label = Label()
label.Text = '结束日期'
label.Location = Point(180, 40)
label.Width = 75
label.TextAlign = ContentAlignment.MiddleCenter
form.Controls.Add(label)

dpicker2 = DateTimePicker()
dpicker2.Tag = 'date2'
dpicker2.Location = Point(260, 40)
dpicker2.Size = Size(90, 50)
dpicker2.Format = DateTimePickerFormat.Custom
dpicker2.CustomFormat = 'yyyy-MM-dd'
dpicker2.Value = DateTime.Now
dpicker2.ValueChanged += dpicker2_ValueChanged
form.Controls.Add(dpicker2)


button = Button()
button.Text = '点 我'
button.Location = Point(155, 85)
form.Controls.Add(button)


def click(sender, event):
    MessageBox.Show('起始日期: {0}, 结束日期: {1}'.format(dpicker1.Value.ToString('yyyy-MM-dd'), dpicker2.Value.ToString('yyyy-MM-dd')), '结果', MessageBoxButtons.OK, MessageBoxIcon.Information)

button.Click += click


def Main():
    Application.Run(form)


if __name__ == '__main__':
    Main()
