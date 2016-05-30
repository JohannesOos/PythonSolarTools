# -*- coding: utf-8 -*-
"""
Created on Mon May 30 06:51:03 2016

@author: Oos
"""

import pandas as pd
import numpy as np
from  xlrd import open_workbook

#loadprofile from FLUKE CSV measurement file
loadprofile = []

#production profile in hourly values
solarProfile = []


from xlrd import open_workbook

book = open_workbook('C:/Users/Oos/sample_Project_HourlyRes_EW.xls')
sheet = book.sheet_by_index(0)

#define column number
columnWithData = 6

#define rows with data
rowsWithData = range(13,8773)


E_Solar = []
for row_index in rowsWithData:
    value = sheet.cell(row_index, columnWithData).value       
    E_Solar.append(value)

if len(E_Solar) == 8760:
    print 'One year hourly read in'
#print E_Load

