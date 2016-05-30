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
else:
    print 'Check length of Solar file'

   
#use this to adjust file
#Option 1 is values in different resolution for the entire year
def x_to_hourly_entire_year(E_Load):
    """ E_Load: List of load value for entire year
    resolution is multiple of one hour """
    goal_length = 8760
    ratio =  goal_length / len(E_Load)
    new_profile = []
    for a in range(len(E_Load)):
        for b in range(ratio):
            new_profile.append(E_Load[a])
    return new_profile
    
#Option2 is if there are not enough values    
def x_to_yearly(E_Load, resolution):
    """
    E_Load: List of production value starting and ending with midnight
    resolution: resolution of values in minutes Must be divisor of 1 hour
    """
    goal_length = 8760 *60/resolution
    ratio =  goal_length / len(E_Load)
    new_profile = []
    for a in range(ratio):
        for b in range(len(E_load)):
            new_profile.append(E_Load[b])
    return new_profile
    



