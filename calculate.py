"""
Description
-------------------------------------------------------------------------------
Contains function do_thermal_electrical_part(...) which does the calculations
of the thermal and electrical part. The heat demand is met by iterating over the 
thermal elements of the system according to the thermal priority given. The electrical
demand is met in a similar way. However, electrical production, when larger than 
consumption is exported to the grid. The hourly values are then entered into the 
excel sheet.

Input
-------------------------------------------------------------------------------
Config module with all the global variables. TRY weather data and thermal and 
electrical load profiles of the building.

Output
-------------------------------------------------------------------------------
Excel sheets with hourly data of the production/consumption of the various elements
present in the  system.
"""

#------------------------------------------------------------------------------
#Importing all the required modules

#To write the results into an excel file
from xlrd import open_workbook
from xlutils.copy import copy
#Contains all the global variables like yearly heat demand,capcaity of CHP,boiler etc
import config

#Import sizing.py for functions and methods related to sizing of the system
import sizing

import annuity
import emissions
import database

#os.chdir("D:/Python_Simulation/CHP")
#Exception for errors
class Error(Exception):
    pass


#------------------------------------------------------------------------------
#Function to do the required calculations.
def do_thermal_electrical_part(workbook_name,worksheet_name,system, modulating_CHP='n',loss='y',hourly_excels='n',reference=False):
    
    #Obtaining the thermal and electrical priorities from the worksheet name
    priority = worksheet_name.split(',')
    th_priority = str(priority[0])
    el_priority = str(priority[1])
    
    #Sizing the system
    sizing.size_system(system)
    config.hours=0
    #Re-initialising the variables to zero
    q_CHP_hourly = [0]*8760
    q_boiler_hourly = [0]*8760
    q_th_storage_hourly = [0]*8760
    q_storage_hourly = [0]*8760
    q_sol_thermal_hourly = [0]*8760
    q_el_heater_hourly = [0]*8760
    sol_thermal_losses_hourly = [0]*8760
    
    #--------------------------------------------------------------------------
    #Calcualtions for the thermal part
    #Meeting the heat consumption according to priority and capacity on hourly basis
    for i in range (0,8760):    #We have 8761 values. Iterating through each value
    
        q_hourly=config.thermaldemand[i]   #The heat consumption is stored in the variable q_hourly
        q_th_storage_hourly[i] = config.dict_q["3"]
        if q_hourly==0:
            if loss in ('yY'):
                config.total_loss += 0.01*config.dict_q["3"]
                config.dict_q["3"] = 0.99*config.dict_q["3"]
            continue
        else:
            
            for th_technology in th_priority:
            
            ##Technologies are iterated over according to their priority
        
        
                #If the technology is CHP and it is not modulating, only max capacity
                #or switched off mode is considered        
                if th_technology=='1' and modulating_CHP == 'n':
                
                    #When hourly consumption is lesser than hourly production and thermal storage is present,
                    #switch on and store the excess heat in thermal storage
                    if q_hourly<=config.dict_cap["1"] and "3" in th_priority:
                
                        #Check for storage availability. If available store excess heat in storage
                        if (config.dict_cap["1"]-q_hourly)<=(config.dict_cap["3"]-config.dict_q["3"]):
                            q_CHP_hourly[i] = config.dict_cap["1"]
                            config.dict_q["1"] += q_CHP_hourly[i]
                            config.dict_q["3"] = config.dict_q["3"]+q_CHP_hourly[i]-q_hourly
                            q_hourly = 0
                            
                            break
            
                            #When thermal storage is not available, do nothing. 
                            #Continue with the next technology            
                
                    #When hourly consumption is greater than hourly production
                    #switch on and contribute
                    elif q_hourly>config.dict_cap["1"]:
                        q_CHP_hourly[i] = config.dict_cap["1"]
                        config.dict_q["1"] += q_CHP_hourly[i]
                        q_hourly -= q_CHP_hourly[i]
                    
                #If CHP is modulating
                if th_technology=='1' and modulating_CHP=='y':
                    q_CHP_hourly_max = config.dict_cap["1"]
                    #When consumption is lesser than 30% of capacity and thermal 
                    # storage is present, check for storage availability and switch
                    # on if available. If not move on to the next technology.
                    if q_hourly<0.3*q_CHP_hourly_max and "3" in th_priority:
                    
                        #Check for storage availability. If available store excess heat in storage
                        if (0.3*q_CHP_hourly_max-q_hourly)<=(config.dict_cap["3"]-config.dict_q["3"]):
                            q_CHP_hourly[i] = 0.3*q_CHP_hourly_max
                            config.dict_q["1"] += q_CHP_hourly[i]
                            config.dict_q["3"] = config.dict_q["3"]+q_CHP_hourly[i]-q_hourly
                            q_hourly = 0
                            break
                
                    #If thermal demand lies between 30% and 100%, contribute to the 
                    # consumption.
                    elif 0.3*q_CHP_hourly_max<=q_hourly and q_hourly<=q_CHP_hourly_max:
                        q_CHP_hourly[i] = q_hourly
                        config.dict_q["1"] += q_CHP_hourly[i]
                        q_hourly = 0
                        break
            
                    #When hourly consumption is lesser than hourly production and thermal storage is not present,
                    #do nothing. Continue with the next technology            
            
            
                    #When hourly consumption is greater than hourly production
                    #switch on and contribute
                    elif q_hourly>q_CHP_hourly_max:
                        q_CHP_hourly[i] = config.dict_cap["1"]
                        config.dict_q["1"] += q_CHP_hourly[i]
                        q_hourly -= q_CHP_hourly[i]
    
                
                #If technology is a peak load device like boiler 
                if th_technology=='2':
                
                    #If hourly consumption is lesser than hourly production, meet the consumption entirely
                    if q_hourly<=config.dict_cap["2"]:
                        q_boiler_hourly[i] = q_hourly
                        config.dict_q["2"] += q_boiler_hourly[i]
                        q_hourly = 0
                        break
                
                    #If hourly consumption is greater than production, contribute and move on to the nexr
                    #technology
                    else:
                        q_boiler_hourly[i] = config.dict_cap['2']
                        q_hourly -= config.dict_cap["2"]
                        config.dict_q["2"] += q_boiler_hourly[i]
                    
                    
                #If technology is a peak load device like electric resistance heater
                if th_technology == '5':
                
                    #If hourly consumption is lesser than hourly production, meet the consumption entirely
                    if q_hourly<=config.dict_cap[th_technology]:
                        q_el_heater_hourly[i] = q_hourly
                        config.dict_q["5"] += q_el_heater_hourly[i]
                        q_hourly = 0
                        break
                
                    #If hourly consumption is greater than production, contribute and move on to the nexr
                    #technology
                    else:
                        q_el_heater_hourly[i] = config.dict_cap['5']
                        config.dict_q["5"] += q_el_heater_hourly[i]
                        q_hourly -= q_el_heater_hourly[i]
                    
                    
                #Some form of thermal storage
                if th_technology == '3':
                    
                    #If heat demand is lesser than the heat stored in the storage, meet it entirely.
                    if q_hourly<=config.dict_q["3"]:
                        q_storage_hourly[i] = q_hourly
                        config.dict_q["3"] -= q_storage_hourly[i]
                        q_hourly = 0
                        break
                
                    #If the heat demand is more than the heat stored in the storage, contribute
                    #and move onto the next technology
                    else:
                        q_storage_hourly[i] = config.dict_q["3"]
                        q_hourly -= q_storage_hourly[i]
                        config.dict_q["3"] = 0
                    
            
                
                #If technology is solar thermal collector. Storage is necessary?      
                if th_technology == '4':
                    #Get hourly heat production from weather data
                    q_sol_thermal_hourly[i]=0.7*config.dict_cap["4"]*config.global_radiation[i]/1000 #In kW
                    if q_hourly<=q_sol_thermal_hourly[i] and "3" in th_priority:
    
                        #Check for storage availability. If available store excess heat in storage
                        if (q_sol_thermal_hourly[i]-q_hourly)<=(config.dict_cap["3"]-config.dict_q["3"]):
                            config.dict_q["3"] += q_sol_thermal_hourly[i]-q_hourly
                            config.dict_q["4"] += q_sol_thermal_hourly[i]
                            q_hourly = 0
                            break
                            
                        #If storage is not available store as much as possible and excess heat is wasted.
                        else:
                            sol_thermal_losses_hourly[i] = q_sol_thermal_hourly[i] - q_hourly - config.dict_cap["3"] + config.dict_q["3"]
                            q_sol_thermal_hourly[i] = q_hourly + config.dict_cap["3"] - config.dict_q["3"]
                            config.dict_q["4"] += q_sol_thermal_hourly[i]                        
                            config.dict_q["3"] = config.dict_cap["3"]
                            q_hourly = 0
                            break
                        
                    #When hourly consumption is lesser than hourly production and thermal storage is not present,
                    #heat is wasted            
                    elif q_hourly<=q_sol_thermal_hourly[i] and "3" not in th_priority:
                        sol_thermal_losses_hourly[i] = q_sol_thermal_hourly[i] - q_hourly
                        q_sol_thermal_hourly[i] = q_hourly
                        config.dict_q["4"] += q_sol_thermal_hourly[i]
                        q_hourly = 0
                
                    #When hourly consumption is greater than hourly production
                    #contribute and move on to the next technology
                    else:
                        q_hourly -= q_sol_thermal_hourly[i]
                        config.dict_q["4"] += q_sol_thermal_hourly[i]
            #At the end of the priotiy list if the heat demand is not being met, raise exception
            if q_hourly != 0:
#                continue
#                if config.flag==0:
#                    print i
#                    config.flag=1
###                print q_CHP_hourly[i-1],q_CHP_hourly[i]
###                print q_hourly, i, config.dict_q["1"],config.dict_q["3"],config.dict_cap["1"],config.dict_cap["3"]
                return
#        
            #If thermal losses are to be considered, 1% of the heat present in the 
            # storage is assumed to wasted
            if loss in ('yY'):
                config.total_loss += 0.01*config.dict_q["3"]
                config.dict_q["3"] = 0.99*config.dict_q["3"]
    

        
            #Counting the number of times the CHP is switched on.
            if q_CHP_hourly[i]!=0 and q_CHP_hourly[i-1]==0:
                config.on_CHP_count += 1
                
            if q_CHP_hourly[i]!=0:
                config.hours += 1

    #--------------------------------------------------------------------------
    #--------------------------------------------------------------------------
    # Assuming everything is exported to the grid

    for el_technology in el_priority:
        if el_technology == "6":
            config.dict_pel_grid["6"] = .15*0.8*config.dict_cap["6"]*config.tot_radiation/1000 #In kW
    
    if hourly_excels == 'y':
        write_hourly_excel(workbook_name,worksheet_name,q_CHP_hourly,\
        q_boiler_hourly,q_storage_hourly,q_sol_thermal_hourly,q_el_heater_hourly,q_th_storage_hourly)
    
    calculate_KPI(system,workbook_name,th_priority,el_priority,reference)
    
def write_hourly_excel(workbook_name,worksheet_name,q_CHP_hourly,\
q_boiler_hourly,q_storage_hourly,q_sol_thermal_hourly,q_el_heater_hourly,q_th_storage_hourly):  

    #--------------------------------------------------------------------------
    #write all values into the worksheet   
#    if worksheet_name == '2,0':
#        continue
#    else:
    workbook = open_workbook(workbook_name)
    idx = workbook.sheet_names().index(worksheet_name)
    workbook = copy(workbook)
    worksheet = workbook.get_sheet(idx)
    worksheet.write(0,0,"Hour")
    worksheet.write(0,1,"Hourly Thermal Demand")
    worksheet.write(0,2,"CHP Production"+str(config.dict_cap["1"]))
    worksheet.write(0,3,"Boiler Production"+str(config.dict_cap["2"]))
    worksheet.write(0,4,"Heat given by Thermal Storage"+str(config.dict_cap["3"]))
    worksheet.write(0,5,"Solar Thermal Collector Production"+str(config.dict_cap["4"]))
    worksheet.write(0,6,"Electrical Resistance Heater Production"+str(config.dict_cap["5"]))
    worksheet.write(0,7,"Thermal Energy present in storage")
           
    for i in range (0,8760):
        worksheet.write(i+1,0,i)
        worksheet.write(i+1,1,config.thermaldemand[i])
        worksheet.write(i+1,2,q_CHP_hourly[i])
        worksheet.write(i+1,3,q_boiler_hourly[i])
        worksheet.write(i+1,4,q_storage_hourly[i])
        worksheet.write(i+1,5,q_sol_thermal_hourly[i])
        worksheet.write(i+1,6,q_el_heater_hourly[i])
        worksheet.write(i+1,7,q_th_storage_hourly[i])      
    workbook.save(workbook_name)

def calculate_self_consumption(hybrid_annuity,boiler_annuity,CHP_heat):
    #print hybrid_annuity,boiler_annuity,CHP_heat
    if CHP_heat==0:
        return 0
    else:
        revenue_needed_for_bp = boiler_annuity-hybrid_annuity
        sc=revenue_needed_for_bp/(.3*.2*CHP_heat*0.142377502727*11.593633866)*0.6*100
        return sc
    
def calculate_KPI(system,workbook_name,th_priority,el_priority,reference):
    
    CHP_annuity = 0
    boiler_annuity=0
    th_storage_annuity = 0
    sol_thermal_annuity = 0
    el_heater_annuity = 0
    PV_annuity = 0
    CHP_a,CHP_CRC,CHP_bonus,CHP_Ank,CHP_Anv,CHP_Anb,CHP_Ane = 0,0,0,0,0,0,0
    boiler_CRC,boiler_Ank,boiler_Anv,boiler_Anb=0,0,0,0
    
    for technology in system:
        if technology == "1":
            if '3' in system:
               CHP_a,CHP_CRC,CHP_bonus,CHP_Ank,CHP_Anv,CHP_Anb,CHP_Ane,CHP_annuity = annuity.get_annuity_CHP(config.dict_cap["1"], config.dict_q["1"],True)
            else:
                CHP_a,CHP_CRC,CHP_bonus,CHP_Ank,CHP_Anv,CHP_Anb,CHP_Ane,CHP_annuity = annuity.get_annuity_CHP(config.dict_cap["1"], config.dict_q["1"],False)
            
        if technology == "2":
            boiler_CRC,boiler_Ank,boiler_Anv,boiler_Anb,boiler_annuity = annuity.get_annuity_boiler(config.dict_cap["2"],config.dict_q["2"],reference)

        if technology == "3":
            th_storage_annuity = annuity.get_annuity_thstorage(config.dict_cap["3"])

        if technology == "4":
            sol_thermal_annuity = annuity.get_annuity_sthermal(config.dict_cap["4"])

        if technology == "5":
            el_heater_annuity = annuity.get_annuity_elheat(config.dict_cap["5"],config.dict_q["5"])

        if technology == "6":
            PV_annuity = annuity.get_annuity_pv(config.dict_cap["6"],config.dict_pel_grid["6"])
    
    Total_annuity = CHP_annuity + boiler_annuity + th_storage_annuity + sol_thermal_annuity + el_heater_annuity + PV_annuity
    
    Total_emissions = 0

    for technology in system:
        if technology == "1":
            Total_emissions += emissions.get_emissions_CHP(config.dict_q["1"])
        if technology == "2":
            Total_emissions += emissions.get_emissions_boilers(config.dict_q["2"])
        if technology == "5":
            Total_emissions += emissions.get_emissions_elheat(config.dict_q["5"])
        if technology == "6":
            Total_emissions += emissions.get_emissions_pv(config.dict_pel_grid["6"])
            
    Total_pef = 0
    Total_pef = (config.dict_q["1"]*0.7 + config.dict_q["2"]*1.1 + config.dict_q["4"]*1 + config.dict_q["5"]*2.8 )/(config.dict_q["1"]+config.dict_q["2"]+config.dict_q["4"]+config.dict_q["5"])
        
    sc_percentage = calculate_self_consumption(Total_annuity, config.reference_annuity, config.dict_q["1"])
    
    config.KPI.append([workbook_name, int(th_priority), int(el_priority), \
    Total_annuity,Total_emissions,Total_pef,config.total_loss, \
    config.dict_cap["1"],config.dict_q["1"],\
    config.on_CHP_count,config.hours,\
    config.dict_cap["2"],config.dict_q["2"],\
    database.change_kwh_to_litres(config.dict_cap["3"]),config.dict_q["3"],\
    config.dict_cap["4"],config.dict_q["4"],\
    config.dict_cap["5"],config.dict_q["5"],\
    config.dict_cap["6"],config.dict_pel_grid["6"],\
    CHP_annuity,\
    boiler_annuity,\
    th_storage_annuity,\
    sol_thermal_annuity,\
    el_heater_annuity,\
    PV_annuity,\
    sc_percentage])
    
    config.Economic_Analysis.append([workbook_name,int(th_priority),Total_annuity,config.dict_cap["1"],config.dict_q["1"],config.dict_q["1"]*.5,CHP_a,CHP_CRC,CHP_bonus,CHP_Ank,CHP_Anv,CHP_Anb,CHP_Ane,CHP_annuity,\
    config.dict_cap["2"],config.dict_q["2"],boiler_CRC,boiler_Ank,boiler_Anv,boiler_Anb,boiler_annuity])
    


#==============================================================================
#==============================================================================
#Testing purposes
#do_thermal_electrical_part('CHP+Thermal Storage.xlsx','13,1','13', modulating_CHP='y',loss='y',hourly_excels='n')
#sizing.size_system("3125")            
#do_thermal_elctrical_part("3125", "15")
#print config.q_yearly 
#print config.on_CHP_count           
#    for i in ('1','2','3','4','5'):     
#        print i,config.dict_cap[i],config.dict_q[i]
#    print ("\n")
#    for i in ('1','6'):
#        print i, config.dict_pel_grid[i]
#    print Total_annuity
#    print Total_emissions
#    print ("\n\n\n\n")
#    
# 
#do_thermal_electrical_part("KPI.xls","312,1","312")
