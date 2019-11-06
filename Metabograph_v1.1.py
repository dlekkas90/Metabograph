# -*- coding: utf-8 -*-
"""
Created on Mon Dec 03 12:00:24 2018

@author: dlekkas
"""

print '\n', 'P R O G R A M  I N I T I A L I Z E D . . .', '\n'*2

print '\t'*4, '+----------- METABOGRAPH v1.1 -----------+'
print '\t'*3, '   < CLAMS Data Parser & Grapher for Circadian Research >', '\n'
print '\t'*3, '    [Conceived & Coded by DAMIEN LEKKAS (c) 2018 - 2019]', '\n'*2

print ' > THE FOLLOWING PROGRAM IMPLEMENTATION WAS DESIGNED FOR USE BY DR. GEORGIOS PASCHOS'
print '   AND ASSOCIATES OF THE FITZGERALD LAB AT THE UNIVERSITY OF PENNSYLVANIA PERELMAN'
print '   SCHOOL OF MEDICINE. EXTERNAL USE OF THIS PROGRAM WITHOUT AUTHOR CONSENT IS STRICTLY'
print '   PROHIBITED.', '\n'

import xlrd
import xlwt
import xlsxwriter
import datetime
from scipy.stats import sem

go = False
while go == False:
    CLAMS_in = raw_input(' > TO BEGIN, PLEASE SPECIFY THE FULL NAME OF THE CLAMS EXCEL OUTPUT FILE TO PARSE: ')
    mice = input(' > PLEASE SPECIFY THE NUMBER OF MICE REFLECTED IN THE DATA: ')

    flag = ''
    while flag != 'y' and flag != 'n':
        flag = raw_input(' > PROCEED WITH DATA PARSING FOR ' + CLAMS_in + ' CONTAINING DATA ON ' + str(mice) + ' MICE (y/n)? ')

    if flag == 'y':
        go = True
        
CLAMS_in_root = CLAMS_in.split(".")[0]
print '\n', ' > CALCULATING AND PARSING . . .', '\n'*2

workbook = xlrd.open_workbook(CLAMS_in)

data_limit = 0 
start_index_lengths = [] 

#MAJOR LOOP TO GRAB DATA FROM EACH SHEET IN WORKBOOK
m = 1
Data_by_mouse = []
for sheet in range(mice):
    worksheet = workbook.sheet_by_index(sheet)    
   
    data_limit_idx = len(worksheet.col_values(0)) - 2
       
    #Parameters to pull
    timeStamps_raw = []
    timeStamps_datetime = []
    ZTs = []
    VO2_list = []
    VCO2_list = []
    RER_list = []
    Heat_list = []
    Food_intake_list = []
    Food_acc_list = []
    Drink_intake_list = []
    Drink_acc_list = []
    XAmb_list = []
    XTOT_list = []
    ZTOT_list = []
    Temp_list = []
    Params_by_ZT = []
    
    for c in range(25, data_limit_idx-1):
        timeStamps_raw.append(float(str((worksheet.cell(c,2))).encode("ascii", "ignore").split(":")[1]))
        timeStamps_datetime.append(datetime.datetime(*xlrd.xldate_as_tuple(timeStamps_raw[-1], workbook.datemode)))
        timeStamps_datetime[-1] = timeStamps_datetime[-1].hour
        
        VO2_list.append(float(str(worksheet.cell(c,3)).encode("ascii", "ignore").split(":")[1]))
        VCO2_list.append(float(str(worksheet.cell(c,8)).encode("ascii", "ignore").split(":")[1]))
        RER_list.append(float(str(worksheet.cell(c,13)).encode("ascii", "ignore").split(":")[1]))
        Heat_list.append(float(str(worksheet.cell(c,14)).encode("ascii", "ignore").split(":")[1]))
        Food_intake_list.append(float(str(worksheet.cell(c,17)).encode("ascii", "ignore").split(":")[1]))
        Food_acc_list.append(float(str(worksheet.cell(c,18)).encode("ascii", "ignore").split(":")[1]))
        Drink_intake_list.append(float(str(worksheet.cell(c,19)).encode("ascii", "ignore").split(":")[1]))
        Drink_acc_list.append(float(str(worksheet.cell(c,20)).encode("ascii", "ignore").split(":")[1]))
        XAmb_list.append(float(str(worksheet.cell(c,22)).encode("ascii", "ignore").split(":")[1]))
        XTOT_list.append(float(str(worksheet.cell(c,21)).encode("ascii", "ignore").split(":")[1]))
        ZTOT_list.append(float(str(worksheet.cell(c,23)).encode("ascii", "ignore").split(":")[1]))
        Temp_list.append(float(str(worksheet.cell(c,26)).encode("ascii", "ignore").split(":")[1]))
        
    for hour in timeStamps_datetime:
        if hour > 6 and hour < 24:
            ZTs.append(hour - 7)
           
        else:
            ZTs.append(hour + 17)     
    
    start_index_list = []
    end_index_list = []

    for value in range(len(ZTs)):
        i = 0
        if value == 0:
            start_index_list.append(value)
        
        elif value == len(ZTs) - 1:  
            if ZTs[value] == ZTs[value-1]:
                end_index_list.append(value+1)
                
            elif ZTs[value] != ZTs[value-1]:    
                start_index_list.append(value)
                end_index_list.append(value+1)
                
        elif ZTs[value] != ZTs[value-1]:
            start_index_list.append(value)
      
            while ZTs[value] == ZTs[value+i] and value+i != len(ZTs) - 1:
                i += 1
            end_index_list.append(value)    
    
    start_index_lengths.append(len(start_index_list))        
    start_index_list = start_index_list[0:len(end_index_list)]       
    
    VO2_avs = []
    VCO2_avs = []
    RER_avs = []
    Heat_avs = []
    Food_intake_avs = []
    Food_acc_avs = []
    Drink_intake_avs = []
    Drink_acc_avs = []
    XAmb_avs = []
    XTOT_avs = []
    ZTOT_avs = []
    Temp_avs = []
    Parameters = ['VO2', 'VCO2', 'RER', 'HEAT', 'FOOD_INTAKE', 'ACC_FOOD', 'DRINK_INTAKE', 'ACC_DRINK', 'XAMB', 'XTOT', 'ZTOT', 'TEMP']
    
    for value in range(len(start_index_list)):
        VO2_avs.append(sum(VO2_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        VCO2_avs.append(sum(VCO2_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        RER_avs.append(sum(RER_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        Heat_avs.append(sum(Heat_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        Food_intake_avs.append(sum(Food_intake_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        Food_acc_avs.append(sum(Food_acc_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        Drink_intake_avs.append(sum(Drink_intake_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        Drink_acc_avs.append(sum(Drink_acc_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        XAmb_avs.append(sum(XAmb_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        XTOT_avs.append(sum(XTOT_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        ZTOT_avs.append(sum(ZTOT_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
        Temp_avs.append(sum(Temp_list[start_index_list[value]:end_index_list[value]])/(end_index_list[value]-start_index_list[value]))
    
    Data_by_mouse.append((m, VO2_avs, VCO2_avs, RER_avs, Heat_avs, Food_intake_avs, Food_acc_avs, Drink_intake_avs, Drink_acc_avs, XAmb_avs, XTOT_avs, ZTOT_avs, Temp_avs))
    m += 1 
    
book = xlwt.Workbook(encoding="ascii")
sh1 = book.add_sheet(Parameters[0])
sh2 = book.add_sheet(Parameters[1])
sh3 = book.add_sheet(Parameters[2])
sh4 = book.add_sheet(Parameters[3])
sh5 = book.add_sheet(Parameters[4])
sh6 = book.add_sheet(Parameters[5])
sh7 = book.add_sheet(Parameters[6])
sh8 = book.add_sheet(Parameters[7])
sh9 = book.add_sheet(Parameters[8])
sh10 = book.add_sheet(Parameters[9])
sh11 = book.add_sheet(Parameters[10])
sh12 = book.add_sheet(Parameters[11])
data_sheets = [sh1, sh2, sh3, sh4, sh5, sh6, sh7, sh8, sh9, sh10, sh11, sh12]
    
for m in range(len(Data_by_mouse)):
    for sheet in range(len(data_sheets)):
        data_sheets[sheet].write(0,m+1, Parameters[sheet] + ' ' + str(m+1))

for sheet in range(len(data_sheets)):        
    for x in range(start_index_lengths[m]):
        data_sheets[sheet].write(x+1,0, ZTs[0]+x)   ###can use ZTs[0] because last mouse will be the last samples taken for that time point, thus latest 
        
        
for sheet in range(len(data_sheets)):
    for m in range(len(Data_by_mouse)):
        j = 0
        for data in range(len(Data_by_mouse[m][j + 1])):
            data_sheets[sheet].write(data+1,m+1, Data_by_mouse[m][sheet+1][data])    
            j += 1
    
book.save(CLAMS_in_root + "_parsed_out.xls")       

print '\n', ' > SUCCESS! DATA HAS BEEN PARSED AND WRITTEN TO XLS STANDARD OUT AS ' + CLAMS_in_root + '_parsed_out.xls' + ' IN YOUR CURRENT WORKING DIRECTORY.', '\n'
raw_input(' > PRESS ENTER TO PROCEED WITH GRAPH GENERATION ')

#-----------------------------DATA GRAPHING-----------------------------------#

print '\n', ' > TO GENERATE APPROPRIATE GRAPHS, PLEASE DESIGNATE UP TO 2 DISTINCT EXPERIMENTAL GROUPS FOR EACH MOUSE AS IT APPEARS IN THE RAW DATA: '

mouse_groupings = []
for mouse in range(len(Data_by_mouse)): 
    mouse_groupings.append(raw_input(' > DEFINE GROUP FOR MOUSE ' + str(mouse+1) + ': '))

count_groupings = len(set(mouse_groupings))

group_names = []
for name in mouse_groupings:
    if name not in group_names:
        group_names.append(name)

print '\n'*2, ' > CALCULATING AND REORGANIZING DATA . . . '


grouping_1_data = []
grouping_2_data = []

#Separate data based on grouping designations
if count_groupings == 2:
    for group in range(len(mouse_groupings)):
        if mouse_groupings[group] == mouse_groupings[0]:
            grouping_1_data.append(Data_by_mouse[group])        
            
            
        elif group != mouse_groupings[0]:
            grouping_2_data.append(Data_by_mouse[group])

    grouping_both_data = [grouping_1_data, grouping_2_data] 
#For Group 1, average each parameter at each time point and calculate the SEM
    VO2_avs_g1 = []
    VCO2_avs_g1 = []
    RER_avs_g1 = []
    Heat_avs_g1 = []
    Food_intake_avs_g1 = []
    Food_acc_avs_g1 = []
    Drink_intake_avs_g1 = []
    Drink_acc_avs_g1 = []
    XAmb_avs_g1 = []
    XTOT_avs_g1 = []
    ZTOT_avs_g1 = []
    Temp_avs_g1 = []
    g1_avs_list = [VO2_avs_g1, VCO2_avs_g1, RER_avs_g1, Heat_avs_g1, Food_intake_avs_g1, Food_acc_avs_g1, Drink_intake_avs_g1, Drink_acc_avs_g1, XAmb_avs_g1, XTOT_avs_g1, ZTOT_avs_g1, Temp_avs_g1]
   
    VO2_sem_g1 = []
    VCO2_sem_g1 = []
    RER_sem_g1 = []
    Heat_sem_g1 = []
    Food_intake_sem_g1 = []
    Food_acc_sem_g1 = []
    Drink_intake_sem_g1 = []
    Drink_acc_sem_g1 = []
    XAmb_sem_g1 = []
    XTOT_sem_g1 = []
    ZTOT_sem_g1 = []
    Temp_sem_g1 = []
    g1_sem_list = [VO2_sem_g1, VCO2_sem_g1, RER_sem_g1, Heat_sem_g1, Food_intake_sem_g1, Food_acc_sem_g1, Drink_intake_sem_g1, Drink_acc_sem_g1, XAmb_sem_g1, XTOT_sem_g1, ZTOT_sem_g1, Temp_sem_g1]
    
    
    group_1_data_lengths = []
    for l in range(len(grouping_1_data)):
        group_1_data_lengths.append(len(grouping_1_data[l][1]))
    
    group_1_data_min_length = min(group_1_data_lengths)
    
    for param in range(1, len(grouping_1_data[0])):
        for value in range(group_1_data_min_length):
            temp_val = 0
            temp_vals_list = []
            for mouse in range(len(grouping_1_data)):                             
                 temp_val += grouping_1_data[mouse][param][value]       
                 temp_vals_list.append(grouping_1_data[mouse][param][value])
            g1_sem_list[param-1].append(sem(temp_vals_list))
            g1_avs_list[param-1].append(float(temp_val/len(grouping_1_data))) 
 
#For Group 2, average each parameter at each time point and calculate the SEM    
    VO2_avs_g2 = []
    VCO2_avs_g2 = []
    RER_avs_g2 = []
    Heat_avs_g2 = []
    Food_intake_avs_g2 = []
    Food_acc_avs_g2 = []
    Drink_intake_avs_g2 = []
    Drink_acc_avs_g2 = []
    XAmb_avs_g2 = []
    XTOT_avs_g2 = []
    ZTOT_avs_g2 = []
    Temp_avs_g2 = []
    g2_avs_list = [VO2_avs_g2, VCO2_avs_g2, RER_avs_g2, Heat_avs_g2, Food_intake_avs_g2, Food_acc_avs_g2, Drink_intake_avs_g2, Drink_acc_avs_g2, XAmb_avs_g2, XTOT_avs_g2, ZTOT_avs_g2, Temp_avs_g2]
    
    VO2_sem_g2 = []
    VCO2_sem_g2 = []
    RER_sem_g2 = []
    Heat_sem_g2 = []
    Food_intake_sem_g2 = []
    Food_acc_sem_g2 = []
    Drink_intake_sem_g2 = []
    Drink_acc_sem_g2 = []
    XAmb_sem_g2 = []
    XTOT_sem_g2 = []
    ZTOT_sem_g2 = []
    Temp_sem_g2 = []
    g2_sem_list = [VO2_sem_g2, VCO2_sem_g2, RER_sem_g2, Heat_sem_g2, Food_intake_sem_g2, Food_acc_sem_g2, Drink_intake_sem_g2, Drink_acc_sem_g2, XAmb_sem_g2, XTOT_sem_g2, ZTOT_sem_g2, Temp_sem_g2]
    
    
    group_2_data_lengths = []
    for l in range(len(grouping_2_data)):
        group_2_data_lengths.append(len(grouping_2_data[l][1]))
    
    group_2_data_min_length = min(group_2_data_lengths)
    
    for param in range(1, len(grouping_2_data[0])):
        for value in range(group_2_data_min_length):
            temp_val = 0
            temp_vals_list = []
            for mouse in range(len(grouping_2_data)):                             
                 temp_val += grouping_2_data[mouse][param][value]       
                 temp_vals_list.append(grouping_2_data[mouse][param][value])
            g2_sem_list[param-1].append(sem(temp_vals_list))
            g2_avs_list[param-1].append(float(temp_val/len(grouping_2_data))) 
            
    master_avs_g1g2 = [g1_avs_list, g2_avs_list]
    master_sems_g1g2 = [g1_sem_list, g2_sem_list]    
                  
intervals = []
ranges = [range(0,53), range(48, 101), range(96, 149), range(144, 197), range(192, 245), range(240, 293)]


#Save independent variables (intervals) as list
for x in range(len(Data_by_mouse[-1][1])):
    intervals.append(ZTs[0]+x)
    
zt_difference = intervals[0]-1
          
workbook = xlsxwriter.Workbook(CLAMS_in_root + '_out_graphs.xlsx')

print '\n', ' > DATA REORGANIZED! '
print '\n'*2, ' > NOW PROCEEDING WITH GRAPH GENERATION . . . '

#MAIN LOOP TO GENERATE FULL INTERVAL CHARTS FOR EACH PARAMETER THAT INCLUDE INDIVIDUAL MICE AND GROUPS
param_counter = 0

for p in range(len(Parameters)):
    
    data_sheets[p] = workbook.add_worksheet(Parameters[p])

    data_sheets[p].write_column('A1', intervals)
    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD']

    for mouse in range(mice):
        data_sheets[p].write_column(columns[mouse]+'1', Data_by_mouse[mouse][param_counter+1])
    
    if count_groupings == 2:
        for group in range(count_groupings):
            data_sheets[p].write_column(columns[group+15]+'1', master_avs_g1g2[group][param_counter])
    
    chart1 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })

    marker_shapes = ['square', 'triangle', 'diamond', 'circle', 'x', 'square', 'triangle', 'circle']
    line_colors = ['black', 'red', 'blue', 'green', 'purple', 'orange', 'magenta', 'yellow']
    ms = 0
    for mouse in range(mice):
        chart1.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[mouse] + '$1:$' + columns[mouse] + '$' + str(len(intervals)),
                           'marker': {'type': marker_shapes[ms],
                                      'size': 5,
                                      'border': {'color': 'black'},
                                      'fill': {'none': True}          
                                      },
                           'line': {'color': line_colors[ms]},
                           'name': '10' + str(1+mouse), 
                           'categories': '=' + Parameters[param_counter] + '!$A$1:$A$' + str(len(intervals))
                           })
        ms += 1
      
#Sub-loop that generates chart for group averages at full interval for each parameter             
    
    if count_groupings == 2:
        chart3 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
        marker_shapes_group = ['square', 'triangle']
        
        msg = 0
        for group in range(count_groupings):
            chart3.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[15+group] + '$1:$' + columns[15+group] + '$' + str(len(intervals)),
                               'y_error_bars':  {
                                       'type':  'custom',
                                       'plus_values': master_sems_g1g2[group][p],
                                       'minus_values': master_sems_g1g2[group][p] ,
                                       },
                               'marker': {'type': marker_shapes_group[msg],
                                      'size': 5,
                                      'border': {'color': 'black'},
                                      'fill': {'none': True}
                                      },
                               'line': { 'color': line_colors[msg]},          
                           'name': group_names[group], 
                           'categories': '=' + Parameters[param_counter] + '!$A$1:$A$' + str(len(intervals))
                           })
            msg += 1    
    
        chart3.set_title({'name': Parameters[param_counter] + ' BY GROUP [0 - ' + str(intervals[-1]) + ']'})
        chart3.set_x_axis({'name': 'Interval', 'min': 0, 'max': intervals[-1], 'major_unit': 12})
        chart3.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
        chart3.set_size({'width': 800, 'height': 550})       
        data_sheets[p].insert_chart('Z1', chart3)           
    
#Sub-loop that generates 2 charts for individuals clustered by group for each parameter at full interval
    
    if count_groupings == 2:
        chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
        chart6 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
       
        for group in range(count_groupings):
            ms = 0
            if group == 0:
                for mouse in range(len(grouping_both_data[0])): 
                    chart5.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[grouping_both_data[0][mouse][0]-1] + '$1:$' + columns[grouping_both_data[0][mouse][0]-1] + '$' + str(len(intervals)),     
                    'marker': {'type': marker_shapes[grouping_1_data[mouse][0]-1],
                                      'size': 5,
                                      'border': {'color': 'black'},
                                      'fill': {'none': True}          
                                      },
                            'line': {'color': line_colors[grouping_1_data[mouse][0]-1]},
                            'name': '10' + str(grouping_both_data[0][mouse][0]), 
                            'categories': '=' + Parameters[param_counter] + '!$A$1:$A$' + str(len(intervals))
                            })
                    ms += 1                       
                
                chart5.set_title({'name': Parameters[param_counter] + ' INDIVIDUAL MICE - GROUP 1: ' + group_names[0] + ' [0 - ' + str(intervals[-1]) + ']'})
                chart5.set_x_axis({'name': 'Interval', 'min': 0, 'max': intervals[-1], 'major_unit': 12, 'interval_tick': 12})
                chart5.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
                chart5.set_size({'width': 800, 'height': 550})
    
                data_sheets[p].insert_chart('AM1', chart5)       
                
            elif group == 1:
                for mouse in range(len(grouping_both_data[1])):
                    chart6.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[grouping_both_data[1][mouse][0]-1] + '$1:$' + columns[grouping_both_data[1][mouse][0]-1] + '$' + str(len(intervals)),
                    'marker': {'type': marker_shapes[grouping_2_data[mouse][0]-1],
                                      'size': 5,
                                      'border': {'color': 'black'},
                                      'fill': {'none': True}          
                                      },
                            'line': {'color': line_colors[grouping_2_data[mouse][0]-1]},
                            'name': '10' + str(grouping_both_data[1][mouse][0]), 
                            'categories': '=' + Parameters[param_counter] + '!$A$1:$A$' + str(len(intervals))
                            })
                    ms += 1                 
        
                chart6.set_title({'name': Parameters[param_counter] + ' INDIVIDUAL MICE - GROUP 2: ' + group_names[1] + ' [0 - ' + str(intervals[-1]) + ']'})
                chart6.set_x_axis({'name': 'Interval', 'min': 0, 'max': intervals[-1], 'major_unit': 12, 'interval_tick': 12})
                chart6.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
                chart6.set_size({'width': 800, 'height': 550})
    
                data_sheets[p].insert_chart('AZ1', chart6)           
        
        
    chart1.set_title({'name': Parameters[param_counter] + ' INDIVIDUAL MICE [0 - ' + str(intervals[-1]) + ']'})
    chart1.set_x_axis({'name': 'Interval', 'min': 0, 'max': intervals[-1], 'major_unit': 12, 'interval_tick': 12})
    chart1.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
    chart1.set_size({'width': 800, 'height': 550})
    
    data_sheets[p].insert_chart('J1', chart1)
   
    param_counter += 1

#MAIN LOOP TO GENERATE INTERMEDIATE INTERVAL CHARTS FOR EACH PARAMETER THAT INCLUDE INDIVIDUAL MICE AND GROUPS
param_counter = 0
for p in range(len(Parameters)):
    chart_coords = ['J', 30]
    range_counter = 0
    
    for i in range(len(ranges)):
        chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
        
        ms = 0
        for mouse in range(mice):
            
            if i == 0:
                chart2.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[mouse] + '$' + str(ranges[i][0]) + ':$' + columns[mouse] + '$' + str(ranges[i][-1]-zt_difference),
                                   'marker': {'type': marker_shapes[ms],
                                              'size': 5,
                                              'border': {'color': 'black'},
                                              'fill': {'none': True}          
                                              },
                                   'line': {'color': line_colors[ms]},
                                   'name': '10' + str(1+mouse), 
                                   'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                   })
            
            else:
                chart2.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[mouse] + '$' + str(ranges[i][0]-zt_difference) + ':$' + columns[mouse] + '$' + str(ranges[i][-1]-zt_difference),
                                   'marker': {'type': marker_shapes[ms],
                                              'size': 5,
                                              'border': {'color': 'black'},
                                              'fill': {'none': True}          
                                              },
                                   'line': {'color': line_colors[ms]},
                                   'name': '10' + str(1+mouse), 
                                   'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]-zt_difference) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                   })            
            ms += 1
         
        chart2.set_title({'name': Parameters[param_counter] + ' INDIVIDUAL MICE [' + str(ranges[i][0]) + ' - ' + str(ranges[i][-1]) + ']'})
        chart2.set_x_axis({'name': 'Interval', 'min': ranges[i][0], 'max': ranges[i][-1], 'major_unit': 12})
        chart2.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
        chart2.set_size({'width': 800, 'height': 550})
        
        data_sheets[p].insert_chart(chart_coords[0] + str(chart_coords[1]), chart2)
        chart_coords[1] += 30
        range_counter += 1        

#Sub-loop that generates charts for group averages at all intermediate time intervals       
    if count_groupings == 2:
        chart_coords = ['Z', 30]
        range_counter = 0
        for i in range(len(ranges)):            
            chart4 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
                   
            msg = 0
            for group in range(count_groupings):  
                if i == 0:
                    chart4.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[15+group] + '$' + str(ranges[i][0]) + ':$' + columns[15+group] + '$' + str(ranges[i][-1]-zt_difference),
                                       'y_error_bars': {
                                               'type': 'custom',
                                               'plus_values': master_sems_g1g2[group][p],
                                               'minus_values': master_sems_g1g2[group][p] ,
                                               },
                                       'marker': {'type': marker_shapes_group[msg],
                                       'size': 5,
                                       'border': {'color': 'black'},
                                       'fill': {'none': True},
                                       },
                                       'line': {'color': line_colors[msg]},           
                                       'name': group_names[group], 
                                       'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                       })
                    
                else:
                    chart4.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[15+group] + '$' + str(ranges[i][0]-zt_difference) + ':$' + columns[15+group] + '$' + str(ranges[i][-1]-zt_difference),
                                       'y_error_bars':  {
                                               'type': 'custom',
                                               'plus_values': master_sems_g1g2[group][p],
                                               'minus_values': master_sems_g1g2[group][p] ,
                                               },
                                       'marker': {'type': marker_shapes_group[msg],
                                       'size': 5,
                                       'border': {'color': 'black'},
                                       'fill': {'none': 'True'},
                                       },
                                       'line': {'color': line_colors[msg]},           
                                       'name': group_names[group], 
                                       'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]-zt_difference) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                       })            
                msg += 1
    
            chart4.set_title({'name': Parameters[param_counter] + ' BY GROUP [' + str(ranges[i][0]) + ' - ' + str(ranges[i][-1]) + ']'})
            chart4.set_x_axis({'name': 'Interval', 'min': ranges[i][0], 'max': ranges[i][-1], 'major_unit': 12})
            chart4.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
            chart4.set_size({'width': 800, 'height': 550})      
            
            data_sheets[p].insert_chart(chart_coords[0] + str(chart_coords[1]), chart4)
            chart_coords[1] += 30
            range_counter += 1
    
    if count_groupings == 2: 
        chart_coords_g1 = ['AM', 30]
        chart_coords_g2 = ['AZ', 30]
        range_counter = 0
        for i in range(len(ranges)):
            chart7 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
            chart8 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_line_with_markers' })
        
            for group in range(count_groupings):
                ms = 0
                if group == 0:
                    if i == 0:
                        for mouse in range(len(grouping_both_data[0])): 
                            chart7.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[grouping_both_data[0][mouse][0]-1] + '$' + str(ranges[i][0]) + ':$' + columns[grouping_both_data[0][mouse][0]-1] + '$' + str(ranges[i][-1]-zt_difference),     
                                               'marker': {'type': marker_shapes[grouping_1_data[mouse][0]-1],
                                                          'size': 5,
                                                          'border': {'color': 'black'},
                                                          'fill': {'none': True}          
                                                      },
                                               'line': {'color': line_colors[grouping_1_data[mouse][0]-1]},
                                               'name': '10' + str(grouping_both_data[0][mouse][0]), 
                                               'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                               })
                            ms += 1        
                    else:
                         for mouse in range(len(grouping_both_data[0])): 
                            chart7.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[grouping_both_data[0][mouse][0]-1] + '$' + str(ranges[i][0]-zt_difference) + ':$' + columns[grouping_both_data[0][mouse][0]-1] + '$' + str(ranges[i][-1]-zt_difference),     
                                               'marker': {'type': marker_shapes[grouping_1_data[mouse][0]-1],
                                                          'size': 5,
                                                          'border': {'color': 'black'},
                                                          'fill': {'none': True}          
                                                      },
                                               'line': {'color': line_colors[grouping_1_data[mouse][0]-1]},
                                               'name': '10' + str(grouping_both_data[0][mouse][0]), 
                                               'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]-zt_difference) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                               })
                            ms += 1           
            
                    chart7.set_title({'name': Parameters[param_counter] + ' INDIVIDUAL MICE - GROUP 1: ' + group_names[0] + ' [' + str(ranges[i][0]) + ' - ' + str(ranges[i][-1]) + ']'})
                    chart7.set_x_axis({'name': 'Interval', 'min': ranges[i][0], 'max': ranges[i][-1], 'major_unit': 12})
                    chart7.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
                    chart7.set_size({'width': 800, 'height': 550})
    
                    data_sheets[p].insert_chart(chart_coords_g1[0] + str(chart_coords_g1[1]), chart7)
                    chart_coords_g1[1] += 30
                   
            
                elif group == 1:
                    if i == 0:
                        for mouse in range(len(grouping_both_data[1])):   
                            chart8.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[grouping_both_data[1][mouse][0]-1] + '$' + str(ranges[i][0]) + ':$' + columns[grouping_both_data[1][mouse][0]-1] + '$' + str(ranges[i][-1]-zt_difference),
                                               'marker': {'type': marker_shapes[grouping_2_data[mouse][0]-1],
                                                          'size': 5,
                                                          'border': {'color': 'black'},
                                                          'fill': {'none': True}          
                                                          },
                                               'line': {'color': line_colors[grouping_2_data[mouse][0]-1]},
                                               'name': '10' + str(grouping_both_data[1][mouse][0]), 
                                               'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                               })
                            ms += 1                 
        
            
                    else:
                        for mouse in range(len(grouping_both_data[1])): 
                            chart8.add_series({'values': '=' + Parameters[param_counter] + '!$' + columns[grouping_both_data[1][mouse][0]-1] + '$' + str(ranges[i][0]-zt_difference) + ':$' + columns[grouping_both_data[1][mouse][0]-1] + '$' + str(ranges[i][-1]-zt_difference),     
                                               'marker': {'type': marker_shapes[grouping_2_data[mouse][0]-1],
                                                          'size': 5,
                                                          'border': {'color': 'black'},
                                                          'fill': {'none': True}          
                                                      },
                                               'line': {'color': line_colors[grouping_2_data[mouse][0]-1]},
                                               'name': '10' + str(grouping_both_data[1][mouse][0]), 
                                               'categories': '=' + Parameters[param_counter] + '!$A$' + str(ranges[i][0]-zt_difference) + ':$A$' + str(ranges[i][-1]-zt_difference)
                                               })
                            ms += 1           
                        
                    chart8.set_title({'name': Parameters[param_counter] + ' INDIVIDUAL MICE - GROUP 2: ' + group_names[1] + ' [' + str(ranges[i][0]) + ' - ' + str(ranges[i][-1]) + ']'})
                    chart8.set_x_axis({'name': 'Interval', 'min': ranges[i][0], 'max': ranges[i][-1], 'major_unit': 12})
                    chart8.set_y_axis({'name': Parameters[param_counter], 'min': 0, 'major_gridlines': {'visible': False}})
                    chart8.set_size({'width': 800, 'height': 550})
    
                    data_sheets[p].insert_chart(chart_coords_g2[0] + str(chart_coords_g2[1]), chart8)
                    chart_coords_g2[1] += 30         
            
            range_counter += 1
            
    param_counter += 1
                
   
workbook.close()

print '\n', ' > SUCCESS! GRAPHS HAVE BEEN WRITTEN TO A NEW .XLSX FILE ' + CLAMS_in_root + '_out_graphs.xlsx' + ' IN YOUR CURRENT WORKING DIRECTORY.', '\n'
raw_input(' > THE PROGRAM HAS NOW FINISHED ITS PROCESSING PIPELINE. PRESS ENTER TO TERMINATE ')






    