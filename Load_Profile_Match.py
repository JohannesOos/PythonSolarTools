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


"""get solar production profile"""
book = open_workbook('C:/Users/Oos/sonstiges/PythonSolarTools/sample_Project_HourlyRes_EW.xls')
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
    print 'One year hourly solar production read in'
else:
    print 'Check length of Solar file'

"""get load profile"""
book_load = open_workbook('C:/Users/Oos/sonstiges/PythonSolarTools/power_logger_FLUKE_sample.xlsx')
sheet_load = book_load.sheet_by_index(0)

#define column number
columnWithLoadData = 3
columnWithLoadTime = 0

#define rows with data
rowsWithLoadData = range(1,8498) 

E_Load_prelim = []
for row_index in rowsWithLoadData:
    value = sheet_load.cell(row_index, columnWithLoadData).value       
    E_Load_prelim.append(value)

if len(E_Load_prelim) == 8760:
    print 'One year hourly load read in'
else:
    print 'Check length and resolution of load file'


   
#use this to adjust files
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
        for b in range(len(E_Load)):
            new_profile.append(E_Load[b])
    return new_profile
    
#Option3 is to adjsut E_Solar resolution
def solar_reso_better(E_Solar, resolution):
    """
    E_Solar: List of hourly solar value for entire year
    resolution: resolution of values in minutes Must be divisor of 1 hour
    """
    goal_length = 8760 *60/resolution
    ratio =  goal_length / len(E_Solar)
    new_profile = []
    for a in range(len(E_Solar)):
        for b in range(ratio):
            new_profile.append(E_Solar[a])
    return new_profile    



