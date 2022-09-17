# -*- coding: utf-8 -*-
"""
Created on Sat Sep 17 09:47:19 2022

@author: takahashi-yosh
"""
import xlwings as xw
import pickle
import sys
import os
epath = sys.argv[1]
hight = float(sys.argv[-1])

#%%
try:
    wb = xw.Book(epath)
    ws = wb.sheets('Output')
    ws.range('B1').value = hight
    output = ws.range('B4:AK11').value

    with open(os.path.join(os.path.dirname(epath), 'data.pkl'), 'wb') as f:
        pickle.dump(output, f, protocol=2)
except:
    with open(os.path.join(os.path.dirname(epath), 'data.pkl'), 'wb') as f:
        pickle.dump(None, f, protocol=2)