#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# 压缩同一目录下同一单元的文件

import os
import re
import sys
import zipfile

nameMap = {}
the_reg = re.compile('(.{2,4}?).*?([789七八九])([上下]).*?(\d{1,2})')
ignore = ('zip')


def findSignature(name: str) -> tuple:
    m = the_reg.match(name)
    if m:
        return (m[1], m[2], m[3], m[4])
    else:
        return ()


def Signature2String(sig: tuple) -> str:
    return f'{sig[0]}版{sig[1]}{sig[2]}U{sig[3]}'


for file in os.listdir('./'):
    if file in __file__:
        continue
    _, ext = os.path.splitext(file)
    if ext in ignore:
        continue
    name = findSignature(file)
    if not name:
        continue
    if name not in nameMap:
        nameMap[name] = []
    nameMap[name].append(file)
for name, files in nameMap.items():
    if len(files) < 2:
        print('unmatch file:{files}')
        continue
    z = zipfile.ZipFile(Signature2String(name)+'.zip',
                        'w', zipfile.ZIP_DEFLATED)
    for file in files:
        z.write(file)
    z.close()
