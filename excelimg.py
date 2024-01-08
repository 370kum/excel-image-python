#!/usr/bin/env python3

# Usage:  excelimg.py <sheetname> <cell> <jpgpath> <jpgheight> <jpgwidth> 

import sys
import openpyxl as px
from openpyxl.styles.borders import Border, Side
import numpy as np
import fileinput
import os
import math
import collections

args = sys.argv

sheetname = args[1]
cell = args[2]
imgpath = args[3]

tmpdir = '/tmp/tmp-'+ str(os.getpid())
tmpfile = tmpdir + '/file.xlsx'
os.makedirs(tmpdir)

# データ取込
stdinFile = open(tmpfile,'wb')
stdinFile.write(sys.stdin.buffer.read())
stdinFile.close()

# ワークブック取込
wb = px.load_workbook(tmpfile)
st = wb[sheetname]

# 画像処理
img = px.drawing.image.Image(imgpath)
img.height = int(args[4])
img.width = int(args[5])
st.add_image(img,cell)

wb.save(tmpfile)

# 出力処理
sys.stdout.buffer.write(open(tmpfile,"rb").read())

# 終了処理
os.remove(tmpfile)
os.rmdir(tmpdir)

sys.exit(0)
