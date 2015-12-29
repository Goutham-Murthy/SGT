"""Combining the thermal and electrical part together. YAhooo!!
Start time- 2-30
"""


"""
Code to simulate the heat and electricity demand semi-dynamically.
Inputs- Weather data, SLP of the house
Output- Annuity of the system, yearly carbon emissions and primary energy factors.
"""
#To write the results into an excel file
from xlrd import open_workbook
from xlutils.copy import copy
import os
#Contains all the global variables like yearly heat demand,capcaity of CHP,boiler etc
import config

#Import sizing.py for functions and methods related to sizing of the system
import sizing

#Exception for errors
class Error(Exception):
    pass

def do_thermal_electrical_part(workbook_name,worksheet_name,system, modulating_CHP='n',loss='n'):
    
    priority = worksheet_name.split(',')
    th_priority = str(priority[0])
    el_priority = str(priority[1])
    
    # Size the system
    sizing.size_system(system)

    # Re-initialising the system
    q_CHP_hourly = [0]*8760
    q_boiler_hourly = [0]*8760
    q_th_storage_hourly = [0]*8760
    q_storage_hourly = [0]*8760
    q_sol_thermal_hourly = [0]*8760
    q_el_heater_hourly = [0]*8760
    sol_thermal_losses_hourly = [0]*8760
    
    #Meeting the heat consumption according to priority and capacity on hourly basis
    for i in range (0,8760):    #We have 8761 values. Iterating through each value
    
        q_hourly=config.thermaldemand[i]   #The heat consumption is stored in the variable q_hourly
            
        for th_technology in th_priority:  #Technologies are iterated over according to their priority
        
                #If the technology is CHP, only max capacity or switched off mode is considered        
            if th_technology=='1' and modulating_CHP == 'n':
                
                #When hourly consumption is lesser than hourly production and thermal storage is present,
                #switch on and store the excess heat in thermal storage
                if q_hourly<=config.dict_cap["1"] and "3" in th_priority:
                
                    #Check for storage availability. If available store excess heat in storage
                    if (q_CHP_hourly[i]-q_hourly)<=(config.dict_cap["3"]-config.dict_q["3"]):
                        q_CHP_hourly[i] = config.dict_cap["1"]
                        config.dict_q["1"] += q_CHP_hourly[i]
                        config.dict_q["3"] = config.dict_q["3"]+q_CHP_hourly[i]-q_hourly
                        q_hourly = 0
                        break
            
                        #When hourly consumption is lesser than hourly production and thermal storage is not present,
                        #do nothing. Continue with the next technology            
            
                #When hourly consumption is greater than hourly production
                #switch on and contribute
                elif q_hourly>config.dict_cap["1"]:
                    q_CHP_hourly[i] = config.dict_cap["1"]
                    config.dict_q["1"] += q_CHP_hourly[i]
                    q_hourly -= q_CHP_hourly[i]
                    
                
            if th_technology=='1' and modulating_CHP=='y':
                q_CHP_hourly_max = config.dict_cap["1"]
                if q_hourly<0.3*q_CHP_hourly_max and "3" in th_priority:
                    #Check for storage availability. If available store excess heat in storage
                    if (0.3*q_CHP_hourly_max-q_hourly)<=(config.dict_cap["3"]-config.dict_q["3"]):
                        q_CHP_hourly[i] = 0.3*q_CHP_hourly_max
                        config.dict_q["1"] += q_CHP_hourly[i]
                        config.dict_q["3"] = config.dict_q["3"]+q_CHP_hourly[i]-q_hourly
                        q_hourly = 0
                        break
                    
                elif 0.3*q_CHP_hourly_max<q_hourly and q_hourly<=q_CHP_hourly_max and "3" in th_priority:
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

                
            #If technology is a peak load device like boiler or electric resistance heater
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
                
                
            #If technology is a peak load device like boiler or electric resistance heater
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
                q_th_storage_hourly[i] = config.dict_q["3"]
        
            #If technology is solar thermal collector. Storage is necessary?      
            if th_technology == '4':
                #Get hourly heat production from weather data
                q_sol_thermal_hourly[i]=0.8*config.dict_cap["4"]*config.global_radiation[i]
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
                
                
                
        if loss in ('yY'):
            config.total_loss += 0.01*config.dict_q["3"]
            config.dict_q["3"] = 0.99*config.dict_q["3"]
        #At the end of the priotiy list if the heat demand is not being met, raise exception
        if q_hourly != 0:
            return
        
        if q_CHP_hourly[i]!=0 and q_CHP_hourly[i-1]==0:
            config.on_CHP_count += 1


###Electrical part starts here
    p_CHP_hourly_house = [0]*8760
    p_CHP_hourly_grid = [0]*8760
    p_PV_hourly_house = [0]*8760
    p_PV_hourly_grid = [0]*8760
    p_battery_hourly = [0]*8760
    p_el_battery_hourly = [0]*8760
    p_grid_hourly = [0]*8760
  # Meeting the electrical production and consumption
 
    for i in range (0,8760):                #We have 8761 values. Iterating through each value
        pel_hourly=config.thermaldemand[i] +  q_el_heater_hourly[i] #The electrical consumption is stored in the variable pel_hourly
        config.electricaldemand.append(pel_hourly)
        for el_technology in el_priority+'8':
        
            # When the technology is CHP, try to consume in-house. Otherwise export it to the grid.       
            if el_technology=='1':
            
                p_CHP_hourly = q_CHP_hourly[i]/3
            
                while p_CHP_hourly!=0:
                    
                    #When hourly electricity production of CHP is greater than consumption
                    if p_CHP_hourly>=pel_hourly:
                        p_CHP_hourly_house[i] = pel_hourly      #Meet the elecrticity demand entirely
                        p_CHP_hourly -= pel_hourly              #Add to the electricity consumed in-house
                        pel_hourly = 0
                
                        #Check if the remaining power can be stored in the battery
                        #Store as much as possible in the battery and export the rest to the grid
                        if p_CHP_hourly>(config.dict_cap["7"]-config.dict_pel_house["7"]) and "7" in el_priority:
                            p_CHP_hourly_house[i] += config.dict_cap["7"]-config.dict_pel_house["7"]
                            p_CHP_hourly -= config.dict_cap["7"]-config.dict_pel_house["7"]
                            config.dict_pel_house["7"] += config.dict_cap["7"]-config.dict_pel_house["7"]
                    
                            #Remaining electricity is exported to the grid
                            if p_CHP_hourly>0:
                                p_CHP_hourly_grid[i] = p_CHP_hourly
                                p_CHP_hourly = 0
                        
                        #If the excess can be completely stored in the battery, store it.
                        elif p_CHP_hourly<=(config.dict_cap["7"]-config.dict_pel_house["7"]) and "7" in el_priority:
                            p_CHP_hourly_house[i] += p_CHP_hourly                        
                            config.dict_pel_house["7"] += p_CHP_hourly
                            p_CHP_hourly = 0
                
                        #If battery is not present, export everything to the grid
                        else:
                            p_CHP_hourly_grid[i] = p_CHP_hourly
                            p_CHP_hourly = 0
                
                    #When production is less than consumption, meet the demand as much as possible
                    #and move on to the next technology
                    else:
                        p_CHP_hourly_house[i] += p_CHP_hourly
                        pel_hourly -= p_CHP_hourly      #Meet the elecrticity demand entirely
                        p_CHP_hourly = 0
                    
                config.dict_pel_house["1"] += p_CHP_hourly_house[i]
                config.dict_pel_grid["1"] += p_CHP_hourly_grid[i]
                p_el_battery_hourly[i] = config.dict_pel_house["7"]

            
            #When PV is present
            if el_technology == '6':
                p_PV_hourly = sizing.get_PV_hourly(i)
            
                while p_PV_hourly != 0:
                
                    #If PV produciton is greater than consumption, meet the consumption and store excess in a battery
                    if p_PV_hourly >= pel_hourly:
                        p_PV_hourly_house[i] = pel_hourly
                        p_PV_hourly -= pel_hourly
                        pel_hourly=0
                    
                        #Check for battery availaibility and if available, store it.                
                        if p_PV_hourly>(config.dict_cap["7"]-config.dict_pel_house["7"]) and "7" in el_priority:
                            p_PV_hourly_house[i] += config.dict_cap["7"]-config.dict_pel_house["7"]
                            p_PV_hourly -= config.dict_cap["7"]-config.dict_pel_house["7"]
                            config.dict_pel_house["7"] += config.dict_cap["7"]-config.dict_pel_house["7"] 
                        
                            #Remaining electricity is exported to the grid
                            if p_PV_hourly>0:
                                p_PV_hourly_grid[i] = p_PV_hourly
                                p_PV_hourly = 0
                        
                        #If battery available, store entire electricity in battery
                        elif p_PV_hourly<=(config.dict_cap["7"]-config.dict_pel_house["7"]) and "7" in el_priority:
                            p_PV_hourly_house[i] += p_PV_hourly
                            config.dict_pel_house["7"] += p_PV_hourly
                            p_PV_hourly = 0
                
                        #If battery is not present, export everything to the grid
                        else:
                            p_PV_hourly_grid[i] += p_PV_hourly
                            p_PV_hourly = 0
                
                    #When production is less than consumption, meet the demand as much as possible
                    #and move on to the next technology
                    else:
                        p_PV_hourly_house[i] += p_PV_hourly
                        pel_hourly -= p_PV_hourly      #Meet the as much of the demand as possible
                        p_PV_hourly = 0

                config.dict_pel_house["6"] += p_PV_hourly_house[i]
                config.dict_pel_grid["6"] += p_PV_hourly_grid[i]
                p_el_battery_hourly[i] = config.dict_pel_house["7"]    
            if el_technology == '7':
            
                if config.dict_pel_house["7"] >= pel_hourly:
                    p_battery_hourly[i] = pel_hourly
                    config.dict_pel_house["7"] -= pel_hourly
                    pel_hourly = 0
                else:
                    p_battery_hourly[i] = config.dict_pel_house["7"]
                    pel_hourly -= config.dict_pel_house["7"]
                    config.dict_pel_house["7"] = 0
                p_el_battery_hourly[i] = config.dict_pel_house["7"]
    
            #If demand is still larger than production, import from the grid.
            if el_technology == '8':            
                if pel_hourly>0:
                    p_grid_hourly[i] = pel_hourly
                    config.dict_pel_house["8"] += pel_hourly
                    pel_hourly = 0
                
    workbook = open_workbook(workbook_name)
    idx = workbook.sheet_names().index(worksheet_name)
    workbook = copy(workbook)
    worksheet = workbook.get_sheet(idx)
    worksheet.write(0,0,"Hour")
    worksheet.write(0,1,"Hourly Thermal Demand")
    worksheet.write(0,2,"CHP Production")
    worksheet.write(0,3,"Boiler Production")
    worksheet.write(0,4,"Heat given by Thermal Storage")
    worksheet.write(0,5,"Solar Thermal Collector Production")
    worksheet.write(0,6,"Electrical Resistance Heater Production")
    worksheet.write(0,7,"Thermal Energy present in storage")
    worksheet.write(0,8,"Hourly Electrical Demand")
    worksheet.write(0,9,"Hourly CHP electricity consumption in-house")
    worksheet.write(0,10,"Hourly CHP electricity export to grid")
    worksheet.write(0,11,"Hourly PV electricity consumption in-house")
    worksheet.write(0,12,"Hourly PV electricity export to grid")
    worksheet.write(0,13,"Battery Contribution")
    worksheet.write(0,14,"Grid Contribution")
    worksheet.write(0,15,"Power in battery")
           
    for i in range (0,8760):
        worksheet.write(i+1,0,i)
        worksheet.write(i+1,1,config.thermaldemand[i])
        worksheet.write(i+1,2,q_CHP_hourly[i])
        worksheet.write(i+1,3,q_boiler_hourly[i])
        worksheet.write(i+1,4,q_storage_hourly[i])
        worksheet.write(i+1,5,q_sol_thermal_hourly[i])
        worksheet.write(i+1,6,q_el_heater_hourly[i])
        worksheet.write(i+1,7,q_th_storage_hourly[i])      
        worksheet.write(i+1,8,config.electricaldemand[i])
        worksheet.write(i+1,9,p_CHP_hourly_house[i])
        worksheet.write(i+1,10,p_CHP_hourly_grid[i])
        worksheet.write(i+1,11,p_PV_hourly_house[i])
        worksheet.write(i+1,12,p_PV_hourly_grid[i])
        worksheet.write(i+1,13,p_battery_hourly[i])
        worksheet.write(i+1,14,p_grid_hourly[i])
        worksheet.write(i+1,15,p_el_battery_hourly[i])
    workbook.save(workbook_name)
    
#sizing.size_system("3125")            
#do_thermal_elctrical_part("3125", "15")
#print config.q_yearly 
#print config.on_CHP_count           
#for i in ('1','2','3','4','5'):     
#    print i,config.dict_cap[i],config.dict_q[i]
    for i in ('1','6','7','8'):
        print i, config.dict_pel_house[i], config.dict_pel_grid[i]