#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
import zipfile

nameMap = {}
the_reg = re.compile('(.{2,4}?).*?([789七八九]).*?([上下]).*?(\d{1,2})')
match = ('mp3','docx')
notMatchList = []


def findSignature(name: str) -> tuple:
    m = the_reg.match(name)
    if m:
        return (m[1], m[2], m[3], m[4])
    else:
        return ()


def Signature2String(sig: tuple) -> str:
    flag = 'U'
    if sig[0] in '外研版':
        flag = 'Module'
    return f'{sig[0]}版{sig[1]}{sig[2]}{flag}{sig[3]}'


for file in os.listdir('./'):
    _, ext = os.path.splitext(file)
    if ext[1:] not in match:
        continue
    name = findSignature(file)
    if not name:
        notMatchList.append(file)
        continue
    if name not in nameMap:
        nameMap[name] = []
    nameMap[name].append(file)
for name, files in nameMap.items():
    if len(files) < 2:
        print(f'没有配对的文件:{files}')
        continue
    z = zipfile.ZipFile(Signature2String(name)+'.zip',
                        'w', zipfile.ZIP_DEFLATED)
    for file in files:
        z.write(file)
    z.close()
print(f'文件名不规范的文件:{notMatchList}')
