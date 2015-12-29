#import weekend
#import config
#import csv #Importing csv module for reading the thermal load profiles
#from itertools import islice #Importing islice to iterate over rows
#import os
#import xlrd
#
#os.chdir('D:/Python_Simulation/')
#config.global_radiation = []
##Opening the TRY data. The delimiter is ;.
#print os.getcwd()
#reader = open('./Wetter_Bottrop_Modelica.csv')
#csv_reader = csv.reader(reader,delimiter='\t')
#
#for row in islice(csv_reader,2,None):
#    config.global_radiation.append(float(row[14])+float(row[15]))
#reader.close()
#config.tot_radiation = sum(config.global_radiation)
#
#
#wb = xlrd.open_workbook('./Heat profiles.xlsx')
#for sheet in wb.sheets():
#    number_of_rows = sheet.nrows
#    number_of_columns = sheet.ncols
#    for col in range(number_of_columns):
#        config.thermaldemand = []
#        config.KPI = []
#        config.building_id = str(sheet.cell(0,col).value)
#        for row in range(1, number_of_rows):
#            value  = (sheet.cell(row,col).value)/1000
#            try:
#                value = float(value)
#            except ValueError:
#                pass
#            finally:
#                config.thermaldemand.append(value)
#        
#os.chdir('D:/Python_Simulation/Output')
#weekend.add_reference_cases(modulating_CHP='n',losses='y',hourly_excels='y')
#weekend.write_KPI_excel(os.getcwd())

import weekend
weekend.generate_cases('D:/Python_Simulation/',modulating_CHP='y',losses='y',hourly_excels='n')