# -*- coding: utf-8 -*-
import math 

CHP_database = {'CHP_mikro_ECO_POWER_1': 2.58 ,
                'CHP_mini_ECO_POWER_3': 8 ,
                'CHP_mini_ECO_POWER_4.7': 12.5 ,
                'CHP_XRGI_9kWel': 20, 
                'CHP_mini_ECO_POWER_20' : 42}

boiler_database = {'Boiler_Vitogas200F_11kW': 11,
                   'Boiler_Vitogas200F_15kW':15,
                   'Boiler_Vitogas200F_18kW':18,
                   'Boiler_Vitogas200F_22kW':22,
                   'Boiler_Vitogas200F_29kW':29,
                   'Boiler_Vitogas200F_35kW':35,
                   'Boiler_Vitogas200F_42kW':42,
                   'Boiler_Vitogas200F_48kW':48,
                   'Boiler_Vitogas200F_60kW':60}
                   
thstorage_database = {'Vaillant_VPS_allSTOR_300':300,
                      'Vaillant_VPS_allSTOR_500':500,
                      'Vaillant_VPS_allSTOR_800':800,
                      'Vaillant_VPS_allSTOR_1000':1000,
                      'Vaillant_VPS_allSTOR_1500':1500,
                      'Vaillant_VPS_allSTOR_1200':2000}
                   
def get_CHP_capacity(required_capacity):
    if required_capacity > 42:
        capacity = 42
        
    max_difference = 200000
    for available_capacity in CHP_database.values():
        difference = abs(available_capacity - required_capacity)
        if difference <= max_difference:
            max_difference = difference
            capacity = available_capacity
    return capacity

def get_boiler_capacity(required_capacity):
    max_difference = 200000
    for available_capacity in boiler_database.values():
        if available_capacity >= required_capacity:
            difference = available_capacity - required_capacity
            if difference <= max_difference:
                max_difference = difference
                capacity = available_capacity
    if required_capacity>60:
        capacity = math.ceil(required_capacity/60.0)*60.0        
    return capacity

def get_thstorage_capacity(required_capacity):
    required_capacity = change_kwh_to_litres(required_capacity)    
    max_difference = 200000
    for available_capacity in thstorage_database.values():
        difference = abs(available_capacity - required_capacity)
        if difference <= max_difference:
            max_difference = difference
            capacity = available_capacity
    capacity = change_litres_to_kwh(capacity)
    if required_capacity>2000:
        capacity = required_capacity
    return capacity


def get_sthermal_capacity(required_capacity):
    capacity = math.floor(required_capacity/2.55)*2.55
    return capacity

def get_pv_capacity(required_capacity):
    capacity = math.floor(required_capacity/1.6434)*1.6434
    return capacity
    
def get_elheater_capacity(required_capacity):
    capacity = math.ceil(required_capacity/100.0)*100
    return capacity

def change_kwh_to_litres(kwh):
    litres = kwh*3600000.0/(4180.0*40.0)
    return litres

def change_litres_to_kwh(litres):
    kwh = litres*4180.0*40.0/3600000.0
    return kwh

