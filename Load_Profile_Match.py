# -*- coding: utf-8 -*-
"""
Created on Mon May 30 06:51:03 2016

@author: Oos
"""

#import pandas as pd
#import numpy as np
from  xlrd import open_workbook
#from scipy import optimize as op
#import matplotlib.mlab as mlab
import matplotlib.pyplot as plt

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
book_load = open_workbook('C:/Users/Oos/sonstiges/PythonSolarTools/power_logger_FLUKE_sample_1day.xlsx')
sheet_load = book_load.sheet_by_index(0)

#define column number
columnWithLoadData = 3
columnWithLoadTime = 0

#define rows with data
rowsWithLoadData = range(1,1441) 

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
                must be smaller or equal to 60 min
    """
#    if resolution > 60:
#        if len(E_Load)/(resolution/60/24)
    if len(E_Load)/(60/resolution) % 24 != 0:
        print 'not ful days in file, check file and resolution'
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
    
"""adjsut code here file specific"""
E_Load = x_to_yearly(E_Load_prelim, 1)
E_Solar = solar_reso_better(E_Solar, 1)
 
#check lengths of both files
if len(E_Load) == len(E_Solar):
    print 'both files same length'
else:
    print 'recheck file lengths'
    print ('Solar file length: ' + str(len(E_Solar)) +
            ' Load file length: ' + str(len(E_Load)))
            
            
            
def compare_profiles(E_Solar, E_Load):
    """
    comapres profiles roughly
    E_Solar: lsit of produciton solar syste,
    E_Load: list of demand client
    returns positive vaue if overproduction of solar yearly sum in kWh
    """
    result = []
    reso = len(E_Solar)/8760
    for i in range(len(E_Solar)):
        value = E_Solar[i] -E_Load[i]
        result.append(value)
    sumOfEnergy = sum(result)
    sumOfEnergy_kWh = sumOfEnergy/1000/reso
    return sumOfEnergy_kWh
        
E_Sum = compare_profiles(E_Solar, E_Load)
if E_Sum >0:
    print 'There is too much yearly solar energy: ' + str(E_Sum) + ' kWh'
else:
    print 'There is too little yearly solar energy: ' + str(E_Sum) + ' kWh'
    
    
def monthly_sum(E_Solar, E_Load):
    """
    comapres profiles in monthly sums
    must be compete years
    E_Solar: lsit of produciton solar system,
    E_Load: list of demand client
    returns list of monhtly values positive vaue if overproduction of solar yearly sum in kWh
    """
    result = []
    reso = len(E_Solar)/8760
    month_len = len(E_Solar)/12
    
    month_sum_list = []
    for i in range(len(E_Solar)):
        value = E_Solar[i] -E_Load[i]
        result.append(value)
    for a in range(12):
        sum_month = sum(result[(a*month_len):((a+1)*month_len)])
        month_sum_list.append(sum_month/1000/reso)
    return month_sum_list
        
Month_Sum = monthly_sum(E_Solar, E_Load)
for a in range(len(Month_Sum)):
    if Month_Sum[a] > 0:
        print 'Month: ' + str(a+1) + ' has ' + str(Month_Sum[a]) + ' kWh overproduction'
    else:
        print 'Month: ' + str(a+1) + ' has ' + str(Month_Sum[a]) + ' kWh underproduction'



def daily_sum(E_Solar, E_Load):
    """
    comapres profiles in daily sums
    must be complete years
    E_Solar: lsit of produciton solar system,
    E_Load: list of demand client
    returns list of monhtly values positive vaue if overproduction of solar yearly sum in kWh
    """
    result = []
    reso = len(E_Solar)/8760
    day_len = len(E_Solar)/365
    
    day_sum_list = []
    for i in range(len(E_Solar)):
        value = E_Solar[i] -E_Load[i]
        result.append(value)
    for a in range(365):
        sum_day = sum(result[(a*day_len):((a+1)*day_len)])
        day_sum_list.append(sum_day/1000/reso)
    return day_sum_list
        
Day_Sum = daily_sum(E_Solar, E_Load)
for a in range(len(Day_Sum)):
    if Day_Sum[a] > 0:
        print 'Day: ' + str(a+1) + ' has ' + str(Day_Sum[a]) + ' kWh overproduction'
    else:
        print 'Day: ' + str(a+1) + ' has ' + str(Day_Sum[a]) + ' kWh underproduction' 
        
def plot_energy_sums(E_Solar, E_Load, time='d'):
    if time == 'd':
        time_list = daily_sum(E_Solar, E_Load)
    elif time == 'm':
        time_list = monthly_sum(E_Solar, E_Load)
    else:
        print 'time must be monthly or daily'
        
    if time == 'd' or time == 'm':
        num_values = len(time_list)
    
        plt.plot(range(1,num_values+1),time_list,  'r')
        plt.xlabel('Day')
        plt.ylabel('kWh: minus is too little prod, plus is too much')
        plt.title(r'kWh surplus and deficit')
        
        # Tweak spacing to prevent clipping of ylabel
        plt.subplots_adjust(left=0.15)
        plt.show()

plot_energy_sums(E_Solar, E_Load)
plot_energy_sums(E_Solar, E_Load, 'm')
plot_energy_sums(E_Solar, E_Load, 'k')
    
    
def adjsut_sizeSolar_to_sum_zero(E_Solar, E_Load):
    """
    retunrs new size in % of solar system to have a yearly sum of zero
    between produced and consumed
    """
    E_Solar_Sum = sum(E_Solar)
    E_Load_Sum = sum(E_Load)
    needed_adjsutment_rel = E_Load_Sum / float(E_Solar_Sum) *100
    return needed_adjsutment_rel
    
print (' to come to yearly sum of zero (theoretical), the solar system needs' +
        'to be adjusted to ' + 
        str(adjsut_sizeSolar_to_sum_zero(E_Solar, E_Load)) + '% of its size')


    
def overproduction_accroding_to_resolution(E_Solar, E_Load):
    """
    returns the actual sums of energy, not the theoretical ones
    """
    yearly_prod = sum(E_Solar)/60.0/1000.0 # in kWH
    yearly_load = sum(E_Load)/60.0 / 1000.0 # in kWh
    from_grid = []
    wasted = []
    for i in range(len(E_Solar)):
        dif = E_Solar[i] - E_Load[i]
        if dif >0:
            from_grid.append(0)
            wasted.append(dif)
        else:
            from_grid.append(-dif)
            wasted.append(0)
    
    return [yearly_prod, yearly_load, sum(from_grid)/60.0/1000.0, sum(wasted)/60.0/1000.0]

a = overproduction_accroding_to_resolution(E_Solar, E_Load)    
print ('actual sums are: \n' + 
        'total yearly production in kWh: ' + str(a[0]) + '\n'
        'total yearly load in kWh: ' + str(a[1]) + '\n'
        'total yearly taken from grid in kWh: ' + str(a[2]) +'\n'
        'total yearly taken wasted solar in kWh: ' + str(a[3]) +'\n'
        'total yearly taken from solar in kWh: ' + str(a[0] - a[3]) )
        
    
            
def theoretical_battery(E_Solar, E_Load, bat_size_kWh = 100, bat_charge_eff = 1, 
                        bat_discharge_eff = 1, bat_start_full = True):
    """
    bith inputs must be yearly
    bat_size is storage size in kWh
    returns list of lists in given resolution
    """
    bat_size = bat_size_kWh *1000
    if bat_start_full:
        bat_level = bat_size
    else:
        bat_level = 0
    from_grid = []
    wasted = []
    too_much = [] # simple comaprison without battery
    bat_level_list = []
    supplied_total = []
    
    for i in range(len(E_Solar)):
        dif = E_Solar[i] - E_Load[i]
        if dif >0:  # if too much solar energy
            from_grid.append(0)  # nothign from grid
            too_much.append(dif)
            supplied_total.append(E_Load[i]) # complete load comes from solar
            if bat_level <= bat_size: #if battery not full
                if dif < (bat_size - bat_level): # all fits in battery
                    wasted.append(0)                   
                    bat_level += dif
                else:   #part fits in battery
                    bat_level = bat_size
                    wasted.append(dif - (bat_size - bat_level))
            else: #battery is full
                wasted.append(dif)
        else: #there is too little solar energy                                  
            wasted.append(0)
            too_much.append(0)
            if bat_level < 0.001: # battery empty
                from_grid.append(-dif)
                supplied_total.append(E_Solar[i]) # all solar is supplied
            elif bat_level >= -dif: #battery can power differnece
                from_grid.append(0)
                bat_level +=dif
                supplied_total.append(E_Load[i]) # all supplied combo PV Bat
            else: # battery can power part of differnce
                from_grid.append(-dif+bat_level)
                supplied_total.append(E_Load[i]+bat_level)
                bat_level = 0
        bat_level_list.append(bat_level)
    
    return [from_grid, wasted, too_much, bat_level_list, supplied_total ]

test = theoretical_battery(E_Solar, E_Load)

if len(test[0]) == len(test[1]) and len(test[2]) == len(test[1]) and len(test[3]) == len(test[1]):
    print "same lengths"
else:
    "something went wrong"
    

def plot_the_bat(E_Solar, E_Load, bat_size_kWh = 100, bat_charge_eff = 1, 
                        bat_discharge_eff = 1, bat_start_full = True, sample1=True,
                        dailyPlot = True):
     result = theoretical_battery(E_Solar, E_Load, bat_size_kWh = bat_size_kWh,
                                  bat_charge_eff = bat_charge_eff , 
                                  bat_discharge_eff = bat_discharge_eff, 
                                  bat_start_full = bat_start_full)
                                  
     grid = result[0]
     wasted = result[1]
     #too_much = result[2]
     bat_level = result[3]
     day = 0
    
     if sample1:  # plot sample
        day = len(E_Solar)/365
        plt.plot(range(day),grid[200*day:201*day],  'r', label = "from grid")
        plt.plot(range(day),wasted[200*day:201*day],  'b', label = "wasted = too much")
        plt.plot(range(day),bat_level[200*day:201*day],  'g', label = "battery level")
        #plt.plot(range(day),too_much[:day],  'r')
        
        plt.xlabel('minute of 19th July')
        plt.ylabel('energy in kWh')
        plt.title(r'kWh surplus (+) and deficit (-)')

        #add legend
        legend = plt.legend()
        
        # Tweak spacing to prevent clipping of ylabel
        plt.subplots_adjust(left=0.15)
        plt.show()
        
     if dailyPlot: # plot daily values
        grid_daily = []
        wasted_daily = []
        too_much_daily = []
        bat_level_daily =  [] 
        day = len(E_Solar)/365
               
        for i in range(365):
            grid_daily.append(sum(result[0][(i*day):((i+1)*day)]))
            wasted_daily.append(sum(result[1][(i*day):((i+1)*day)]))
            too_much_daily.append(sum(result[2][(i*day):((i+1)*day)]))
            bat_level_daily.append(sum(result[3][(i*day):((i+1)*day)]))
        
        red = plt.plot(range(365),grid_daily,  'r', label = "grid_daily")
        blue = plt.plot(range(365),wasted_daily,  'b', label = "wasted_daily")
        green = plt.plot(range(365),bat_level_daily,  'g', label = "bat_daily")
        #plt.plot(range(day),too_much[:day],  'r')
        
        plt.xlabel('Day of year')
        plt.ylabel('energy in kWh sum per day')
        plt.title(r'energy daily sums for differnt values in kWh')
        
        #add legend
        plt.legend()
        
        # Tweak spacing to prevent clipping of ylabel
        plt.subplots_adjust(left=0.15)
        plt.show()  
        
     if True: # plot daily values for grid and solar_prod and solar_use
        grid_daily = []
        E_used_daily = []
        E_Solar_prod_Daily =  [] 
        day = len(E_Solar)/365
        
        #to be adjsuted
        for i in range(365):
            grid_daily.append(sum(result[0][(i*day):((i+1)*day)]))
            E_used_daily.append(sum(result[4][(i*day):((i+1)*day)]))
            E_Solar_prod_Daily.append(sum(E_Solar[(i*day):((i+1)*day)]))
        
        plt.plot(range(365),grid_daily,  'r', label = "grid_daily")
        plt.plot(range(365),E_used_daily,  'b', label = "used_daily")
        plt.plot(range(365),E_Solar_prod_Daily,  'g', label = "possible_daily")
        
        plt.xlabel('Day of year')
        plt.ylabel('energy sum per day in kWh')
        plt.title(r'kWh daily sums for differnt values')
        
        #add legend
        plt.legend()
        
        # Tweak spacing to prevent clipping of ylabel
        plt.subplots_adjust(left=0.15)
        plt.show()  

test_plot_level = plot_the_bat(E_Solar, E_Load)   
                
            


    
    
    
    
    


