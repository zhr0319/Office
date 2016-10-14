# -*- coding: utf-8 -*-
# 蒙提霍尔问题
from __future__ import unicode_literals, division
from random import choice
import argparse

parser = argparse.ArgumentParser(description='参数测试')

parser.add_argument('-t', '--times', help='运行次数', default=100)
parser.add_argument('-c', '--change', help='填写y或Y表示改变、n或N表示不改变', choices=['y', 'n', 'Y', 'N'], default='y')

args = parser.parse_args()

Doors = set('ABC')

result = []
change = True if args.change.lower() == 'y' else False
for i in range(int(args.times)):
    car = choice(list(Doors))
    you = choice(list(Doors))
    goat = choice(list(Doors - set([car, you])))
    you = choice(list(Doors - set([goat, you]))) if change else you
    result.append(True if car == you else False)

print '{0:%}'.format(result.count(True) / len(result))
