
"""
Description
-------------------------------------------------------------------------------
Sizing of the various technologies is done here. It needs the
thermal load profile of the building as input. As output, depending on which
technologies are present in the system, it sizes the system and updates the
values of their capacities in the config file. More specifically it updates
config.dict_cap. It also feeds values into config.thermaldemand which is used
in further calculations.

Logic
------------------------------------------------------------------------------- 
The logic followed for the various technologies, if they are present in the
system is as follows-
1) CHP
Maximum rectangle method is followed to dimension the CHP in the presence of 
a peak load device like a boiler or an electric resistance heater. In the 
absence of any peak load device, the CHp is sized according to the peak thermal
demand of the building(s).
2) Boiler and Electric Resistance Heater
The boiler/electric resistance heater, when present in the system, are sized 
according to the peak thermal demand of the building(s).
3) Thermal Storage
Thermal storage is sized depending on the ratio of the peak thermal demand and 
the base load. Refer --- add reference here
4) Solar Thermal Flat Plate Collectors or PV installtion
Depending on the EST, the area available for the installation of a solar
thermal flat plate collector or a PV module is given.
5) Battery
The battery is sized according to one-third of the maximum electrical demand of
the house.

Input
-------------------------------------------------------------------------------
As input the thermal and electrical load profiles of the building is required
in csv format. Also, TRY Weather Data is required in case a solar thermal 
collector or a PV installation is present.

Output
-------------------------------------------------------------------------------
Updates config.dict_cap and config.thermaldemand. If solar thermal or a PV 
installation is present it also feeds values into config.global_radiation.

"""

##-----------------------------------------------------------------------------
#Importing required modules
import config #importing the global variables 
import database

##-----------------------------------------------------------------------------
#Defining function for sizing the given system.
def size_system(system,reference=False):
    
    #re-initialising the config variables
    config.q_yearly=0
    config.dict_q = {'1':0,'2': 0 ,'3':0,'4':0, '5':0,'6':0}
    config.dict_cap = {'1':0,'2': 0 ,'3':0,'4':0, '5':0,'6':0}
    config.dict_pel_grid = {'1':0,'6': 0,'8':0}    
    config.total_loss = 0
    config.on_CHP_count = 0

    #Sort thermal demand in decreasing order for the load distribution curve
    thermaldemand = sorted(config.thermaldemand, reverse=True)
#    config.q_yearly = trapz(thermaldemand, dx=1)
    
    #Finding the maximum rectangle
    maxr=0
    config.hours=0
    for k in range(0,8760):
        config.q_yearly += thermaldemand[k]
        if k*thermaldemand[k]>maxr:
            maxr=k*thermaldemand[k]
            config.hours=k
    print config.hours,thermaldemand[config.hours],thermaldemand[0]
    
    if reference:
        return database.get_boiler_capacity(thermaldemand[0])

#------------------------------------------------------------------------------
#CHP    
    #If CHP is present, it will check for a peak load device. If peak load 
    #device is present, CHP is sized according to maximum rectangle method. 
    # Otherwise it is sized according to peak thermal load.
    if '1' in system:
        if '2' in system or '5' in system:
            config.dict_cap["1"] = database.get_CHP_capacity(thermaldemand[config.hours])
        else:
            config.dict_cap["1"] = database.get_CHP_capacity(thermaldemand[0])

#------------------------------------------------------------------------------
#Boiler    
    #If boiler is present, dimension it to peak thermal demand else capacity 
    #is 0    
    if '2' in system:
        config.dict_cap["2"] = database.get_boiler_capacity(thermaldemand[0])
            
        
#------------------------------------------------------------------------------
#Electric Resistance Heater
    #If electric heater is present, dimension it to peak thermal demand else 
    #capacity is 0    
    if '5' in system:
        config.dict_cap["5"] = database.get_elheater_capacity(thermaldemand[0])
#------------------------------------------------------------------------------
#Thermal Storage    
    #If storage is present dimesion it according to ratio of base load to peak 
    #load. Otherwise it is 0       
    if '3' in system:
        config.dict_cap["3"] = database.get_thstorage_capacity(3*config.dict_cap["1"]) #in kWh
       # config.dict_cap["3"] = config.dict_cap["3"]*3600000/(1000*4.186*80) #in liters
        

#------------------------------------------------------------------------------
#Solar Thermal Collector
    #It is sized depending on the EST the building belongs to
    if '4' in system:
        if '6' in system:
            config.dict_cap["4"] = database.get_sthermal_capacity(5)
        else:
            config.dict_cap["4"] = database.get_sthermal_capacity(10)
    
#------------------------------------------------------------------------------
#Photovoltaics
    #It is sized depending on the EST the building bellongs to
    if '6' in system:
        if '4' in system:
            config.dict_cap["6"] = database.get_pv_capacity(5)
        else:
            config.dict_cap["6"] = database.get_pv_capacity(10)
    


