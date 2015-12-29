# -*- coding: utf-8 -*-
"""
Description
-------------------------------------------------------------------------------
Generates possible scenarios with different systems and priorities. For the 
different scenarios, thermal and electrical calculations are done. Furthermore,
an estimation regarding the annuity, emissions and the primary energy factors
for each of the scenarios is done.

Input
-------------------------------------------------------------------------------
TRY weather data and the thermal and electrical load profiles of the building in
CSV format.

Output
-------------------------------------------------------------------------------
In the output folder, depending on the number of technologies in the system a
folder is created. Inside these folders, the excel files depict the technologies
present in the system. Each worksheet in the excel file represnts different 
priorities for the thermal and electrical part. Each worksheet contains hourly
values of heat/electricity consumption/production of the teachnologies present
in the system.

Output Folder  -with all the results
    |   
    |__> No. of technologies Folder 
        (Ex- 1technologies, 2technologies etc)
            |            
            |____> Excel Sheets with names containing the technologies present
                    (Ex- CHP+Boiler+.xls for an excel file containing data of 
                    the CHP + Boiler system)
                        |
                        |_____> Worksheets with names representing the priorities
                                (Ex- 12,1. Here 12 is the thermal priority and 
                                1 is the electrical priority)
"""   

#------------------------------------------------------------------------------
#Importing required modules
import os
import itertools
import xlsxwriter
import config
import shutil
import xlrd
import calculate
import annuity
import sizing
import csv #Importing csv module for reading the thermal load profiles
from itertools import islice #Importing islice to iterate over rows

def generate_cases(location,modulating_CHP="y",losses="y",hourly_excels="n"):
    if not os.path.exists(location+'/Output'):
        os.makedirs(location+'/Output')
    
    os.chdir(location)
    for the_file in os.listdir(location+'/Output'):
        file_path = os.path.join(location+'/Output', the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception, e:
            print e

    
    config.global_radiation = []
    #Opening the TRY data. The delimiter is ;.
    reader = open('./Wetter_Bottrop_Modelica.csv')
    csv_reader = csv.reader(reader,delimiter='\t')
    
    for row in islice(csv_reader,2,None):
        config.global_radiation.append(float(row[14])+float(row[15]))
    reader.close()
    config.tot_radiation = sum(config.global_radiation)
    
    
    wb = xlrd.open_workbook('./Heat profiles.xlsx')
    os.chdir('./Output/')
    
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        for col in range(number_of_columns):
            config.thermaldemand = []
            config.KPI = []
            config.Economic_Analysis = []
            config.building_id = str(sheet.cell(0,col).value)
            for row in range(1, number_of_rows):
                value  = (sheet.cell(row,col).value)/1000
                try:
                    value = float(value)
                except ValueError:
                    pass
                finally:
                    config.thermaldemand.append(value)
                    
            no,no1,no2,no3,config.reference_annuity=annuity.get_annuity_boiler(sizing.size_system('2',True),sum(config.thermaldemand),True)
            if not os.path.exists(config.building_id):
                os.makedirs(config.building_id)
            os.chdir('./'+config.building_id)
            
#            #------------------------------------------------------------------------------
#            #Delete all folders with the same name (may be present due to old runs)
#            #to avoid confusion
#            for number in range (1,8):
#                if os.path.exists('.'+str(number)+'technologies'):
#                    shutil.rmtree('.'+str(number)+'technologies')
#                if os.path.exists ('./KPI.xls'):
#                    os.remove('./KPI.xls')
            
    
         
            #------------------------------------------------------------------------------
            #Initialising the variables.
            #Thermal technologies are-
            # 1- CHP
            # 2- Boiler
            # 3- Thermal Storage
            # 4- Solar Thermal Collector
            # 5- Electric Resistance Heater
            
            #Electrical technologies apart from CHP are-
            #6- Photovoltaics
                  
            technologies = ['1','2','3','4','5','6']
            th_technologies = ['1','2','3','4','5']
            el_technologies = ['1','6'] 
            
            #------------------------------------------------------------------------------
            #To limit the number of technologies based on thermal demand        
            max_el_technologies = 2
            max_th_technologies = 5
            min_el_technologies = 0
            min_th_technologies = 0
            count = 0
            
            #------------------------------------------------------------------------------
            #Iterating from 1 to the maximum number of technologies 
            for number in range(1,len(technologies)+1):
                
                #Make seperate directories depending on the number of technologies present
                #in the system
                if (number <= max_el_technologies + max_th_technologies-1 \
                and number >= min_el_technologies+min_th_technologies):      
                    if not os.path.exists(str(number)+'technologies'):
                        os.makedirs(str(number)+'technologies')
            
                    #Changing to the required directory in output directory
                    os.chdir('./'+str(number)+'technologies')
                    #----------------------------------------------------------------------
                    #Iterating through the various possible combinations.
                    for combo in itertools.combinations(technologies,number):
                        
            
                        #Proceed further only if there is CHP present in the system
                        if '1' in combo:
                            
                            if '2' in combo and '5' in combo:
                                continue
                            
                            #Proceed further is the number of technologies in the system are
                            # suitable
                            if (len(set(combo) & set(el_technologies)) <= max_el_technologies\
                            and len(set(combo) & set(th_technologies)) <= max_th_technologies \
                            and len(set(combo) & set(el_technologies)) >= min_el_technologies \
                            and len(set(combo) & set(th_technologies)) >= min_th_technologies):               
                                
                                #Workbook has the workbook name
                                #System has the system in terms of numbers. Ex- 12 for 
                                #CHP-Boiler system                    
                                workbook = ""
                                system = ""
                                for i in combo:
                                    system += i
                                    workbook += config.dict1[i] + "+"
                                    
                                #Create excel with the corresponding name
                                excel = xlsxwriter.Workbook(workbook+".xls")
                                
                                #----------------------------------------------------------
                                #Generating the electrical system.
                                el_priority = ""
                                for elements in (set(combo) & set(el_technologies)):
                                    el_priority += elements
                                    
                                #------------------------------------------------------
                                #Generating different thermal priorities for each electrical
                                # priority and iterating through them
                                for th_order in itertools.permutations(set(combo) & set(th_technologies)):
                                        
                                    #print th_order
                                    #Eliminating cases where the boiler or electric 
                                    # resistance heater preceeds the CHP in thermal
                                    # priority                            
                                    if '2' in th_order and th_order.index('2')<th_order.index('1'):
                                        continue
                                    if '5' in th_order and th_order.index('5')<th_order.index('1'):
                                        continue
#                                    if '2' in th_order and '5' in th_order:
#                                        continue
                                        
                                    th_priority = ""
                                    for elements in th_order:
                                        th_priority += elements
                                            
                                    #For each thermal
                                    # priority a seperate worksheet is created.
                                    worksheet = excel.add_worksheet(th_priority+','+el_priority)
                                    count += 1
                                excel.close()
                            
                                #Open the workbook and do the calculations and update all 
                                # the values.
                                excel = xlrd.open_workbook(workbook+".xls")
                                for worksheet in excel.sheet_names():
                                    #print worksheet
                                    calculate.do_thermal_electrical_part(workbook+".xls",worksheet,system,modulating_CHP,losses,hourly_excels,False)
                                print workbook
                                print "============================================"
                                print "++++++++++++++++++++++++++++++++++++++++++"
                #Change directory to output directory
                os.chdir("../")
                
            add_reference_cases(modulating_CHP,losses,hourly_excels)
            #Change directory to output directory
            write_KPI_excel(location+'/Output/'+config.building_id+'/')
            
            if hourly_excels=="n":
                delete_empty_excels()
            
            os.chdir(location+'/Output')
            
def add_reference_cases(modulating_CHP,losses,hourly_excels):
    workbook = 'Reference_case-Boiler'
    excel = xlsxwriter.Workbook(workbook+".xls")
    excel.add_worksheet('2,0')
#    excel.add_worksheet('13,1')
#    excel.add_worksheet('312,1')
#    excel.add_worksheet('12,1')
    excel.close()
    excel = xlrd.open_workbook(workbook+".xls")
    
    calculate.do_thermal_electrical_part(workbook+".xls",'2,0','2',modulating_CHP,losses,hourly_excels,True)
#    calculate.do_thermal_electrical_part(workbook+".xls",'13,1','13',modulating_CHP,losses,hourly_excels)
#    calculate.do_thermal_electrical_part(workbook+".xls",'312,1','312',modulating_CHP,losses,hourly_excels)
#    calculate.do_thermal_electrical_part(workbook+".xls",'12,1','12',modulating_CHP,losses,hourly_excels)
    if hourly_excels == 'n':
        os.remove('Reference_case-Boiler.xls')
    #write_KPI_excel(os.getcwd())

    
def write_KPI_excel(location):
    os.chdir(location)               
    excel = xlsxwriter.Workbook(config.building_id+".xls")  
    worksheet = excel.add_worksheet("Thermal Profile")
    for row in range(0,len(config.thermaldemand)):
        worksheet.write(row,0,row)
        worksheet.write(row,1,config.thermaldemand[row])
    
    # Create a new Chart object.
    chart = excel.add_chart({'type': 'line'})
    
    chart.set_x_axis({
    'name': 'Hours',
    'name_font': {'size': 10, 'bold': True},
    'label_position': 'low'
    })

    chart.set_y_axis({
    'name': 'Thermal Demand in kWh',
    'name_font': {'size': 10, 'bold': True}
    })
        
    chart.set_title({'name': 'Thermal Load Profile'})
        
        
    # Configure the chart.
    chart.add_series({'values':['Thermal Profile',0,1,8760,1],
                      'categories':['Thermal Profile',0,0,8760,0]})
    chart.set_legend({'none': True})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('D4', chart)
                  
    worksheet = excel.add_worksheet("KPI")
    worksheet.write(0,0,"System")
    worksheet.write(0,1,"Thermal Priority")
    worksheet.write(0,2,"Electrical Priority")
    worksheet.write(0,3,"Annuity (Euros)")
    worksheet.write(0,4,"Emissions (kg of CO2)")
    worksheet.write(0,5,"Primary Energy Factor")
    worksheet.write(0,6,"Total Losses (kWh)")
    worksheet.write(0,7,"CHP capacity (kW)")
    worksheet.write(0,8,"CHP heat (kWh)")
    worksheet.write(0,9,"CHP_On_Count")
    worksheet.write(0,10,"CHP_Hours (hours)")                       
    worksheet.write(0,11,"Boiler capacity (kW)")
    worksheet.write(0,12,"Boiler heat (kWh)")
    worksheet.write(0,13,"Storage capacity (liters)")
    worksheet.write(0,14,"Storage heat (kWh)")
    worksheet.write(0,15,"Solar Thermal Are (m2)")
    worksheet.write(0,16,"Solar Thermal heat (kWh)")
    worksheet.write(0,17,"El heater capacity (kW)")
    worksheet.write(0,18,"El heater heat (kWh)")
    worksheet.write(0,19,"PV area (m2)")
    worksheet.write(0,20,"PV El (kWh)")
    worksheet.write(0,21,"CHP Annuity(Euros)")
    worksheet.write(0,22,"Boiler Annuity(Euros)")
    worksheet.write(0,23,"Th Storage Annuity(Euros)")
    worksheet.write(0,24,"Sol Thermal Annuity(Euros)")
    worksheet.write(0,25,"El Heater Annuity(Euros)")
    worksheet.write(0,26,"PV Annuity(Euros)")
    worksheet.write(0,27,"Self Consumption needed for break-even with boiler(%)")
    row = 1
    for item in config.KPI:
        for column in range(0,28):
            worksheet.write(row,column,item[column])
            column +=1
        row += 1
        
    # Create a new Chart object.
    chart = excel.add_chart({'type': 'scatter'})
    
    chart.set_x_axis({
    'name': 'Yearly Emissions',
    'name_font': {'size': 10, 'bold': True},
    'label_position': 'low'
    })

    chart.set_y_axis({
    'name': 'Annuity',
    'name_font': {'size': 10, 'bold': True}
    })
        
    chart.set_title({'name': 'Results'})
        
        
    # Configure the chart. In simplest case we add one or more data series.
    chart.add_series({'values':['KPI',1,3,row,3],
                      'categories':['KPI',1,4,row,4]})
    chart.set_legend({'none': True})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('V4', chart)
    
    worksheet = excel.add_worksheet("Economic Factors in Detail")
    worksheet.write(0,0,"System")
    worksheet.write(0,1,"Thermal Priority")   
    worksheet.write(0,2,"Total Annuity(Euros)")
    worksheet.write(0,3,"Th Capacity of CHP(kW)")
    worksheet.write(0,4,"Heat generated by the CHP(kWh)")
    worksheet.write(0,5,"Electricity generated by the CHP(kWh)")
    worksheet.write(0,6,"Annuity factor")
    worksheet.write(0,7,"CHP Capital Costs(Euros)")
    worksheet.write(0,8,"CHP bonus(Euros)")
    worksheet.write(0,9,"CHP Capital related Annuity(Euros)")
    worksheet.write(0,10,"CHP demand related Annuity(Euros)")
    worksheet.write(0,11,"CHP operation related Annuity(Euros)")
    worksheet.write(0,12,"CHP proceeds related Annuity(Euros)")
    worksheet.write(0,13,"CHP Total Annuity(Euros)")
    worksheet.write(0,14,"Boiler Capacity(kW)")
    worksheet.write(0,15,"Heat generated by the Boiler(kWh)")                       
    worksheet.write(0,16,"Boiler Capital related Costs(Euros)")
    worksheet.write(0,17,"Boiler Capital related Annuity(Euros)")
    worksheet.write(0,18,"Boiler demand related Annuity(Euros)")
    worksheet.write(0,19,"Boiler operation related Annuity(Euros)")
    worksheet.write(0,20,"Total Boielr Annuity(Euros)")
    row = 1
    for item in config.Economic_Analysis:
        for column in range(0,21):
            worksheet.write(row,column,item[column])
            column +=1
        row += 1
    
    excel.close()     
    return
    
def delete_empty_excels():
    if os.path.exists('./1technologies'):
        shutil.rmtree('./1technologies')
    if os.path.exists('./2technologies'):
        shutil.rmtree('./2technologies')
    if os.path.exists('./3technologies'):
        shutil.rmtree('./3technologies')
    if os.path.exists('./4technologies'):
        shutil.rmtree('./4technologies')
    if os.path.exists('./5technologies'):
        shutil.rmtree('./5technologies')
    if os.path.exists('./6technologies'):
        shutil.rmtree('./6technologies')            
