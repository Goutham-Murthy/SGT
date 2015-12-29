import math
import database
#annuity factor as defined by VDI 2067
def get_annuity_factor(q,obperiod):
    a = (q-1)/(1-q**(-obperiod))
    return a
    
#For getting the anuuity    
def get_ank(A0,r,q,obperiod,deperiod):
    a = get_annuity_factor(q,obperiod)
    A = [0] 
    n = int(math.floor(obperiod/deperiod))
    for i in range (1,n+1):
        A.append(A0*(r**(i*deperiod))/(q**(i*deperiod)))
    An = A0
    for i in range (0,n+1):
        An = An + A[i] 
    Rw = A0*r**(n*deperiod)*((n+1)*deperiod-obperiod)/(deperiod*q**obperiod)
    Ank = (An-Rw)*a
    return Ank

def get_b(r,q,obperiod):
    b = (1-(r/q)**obperiod)/(q-r)
    return b
#For getting the annuity of the boiler
#obperiod - observation period
#q - interest-rate factor
#r - price change factor
#deperiod- depreciation period
#bv - price-dynamic cash value factor for consumption-related costs 
#bb - price-dynamic cash value factor for operation-related costs
#bi -  price-dynamic cash value factor for maintenance
#finst - effort for annual repairs as percentage of initial investment
#fwins - effort for annual maintenance and inspection as percentage of initial investment
#n - number of replacements procured within the observation period 
#Rw residual value







def get_annuity_boiler(capacity,Qyearly,reference=False):
    obperiod = 10 #VDI2067
#The following factors are assumed. For boilers and other technologies
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b

#The deprecition period, finst,fwins and effop are different for different capacitites according to 
#VDI 2067    
    if capacity < 100:
        deperiod = 18
        finst = 0.015
        fwins = 0.015
        effop = 10
    elif capacity in range(100,200):
        deperiod = 20
        finst = 0.01
        fwins = 0.015
        effop = 20
    else:
        deperiod = 20
        finst = 0.01
        fwins = 0.02
        effop = 20
    
    a = get_annuity_factor(q,obperiod)
    #Capital related costs for the boiler include price of purchase and installation costs
    if reference:
        CRC = 79.061*capacity + 3229.8
    else:
        CRC = 79.061*capacity + 1229.8
    Ank = get_ank(CRC,r,q,obperiod,deperiod)
    
    #Demand related costs include price of fuel to generate the required heat
    DRC = 0.067*Qyearly/.93
    Anv = DRC*a*bv
    
    #Operation related costs include maintanance and repair
    ORC = 30*effop
    Ain = CRC*(finst+fwins)
    Anb = ORC*a*bb + Ain*a*bi

    #Other costs
    Ans = 0
    
    #Proceeds
    Ane = 0
#    print 'boiler'
#    print Ank
#    print Anv
#    print Anb
#    print Ane
    A = Ane - (Ank+Anv+Anb+Ans)
    
    return CRC,Ank,Anv,Anb,A 

def get_annuity_sthermal(area):
    obperiod = 10
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b
    deperiod = 20
    finst = 0.005
    fwins = 0.01   
    effop = 5
    
    a = get_annuity_factor(q,obperiod)
    #Capital related costs for the solar thermal collector include price of purchase and installation costs
    CRC = 442.8*area + 1000
    Ank = get_ank(CRC,r,q,obperiod,deperiod)

    #Demand related costs include price of fuel to generate the required heat
    DRC = 0
    Anv = DRC*a*bv

    #Operation related costs include maintanance and repair
    ORC = 30*effop
    Ain = CRC*(finst+fwins)
    Anb = ORC*a*bb + Ain*a*bi
    
    #Other costs
    Ans = 0
    
    #Proceeds
    Ane = 0
    
    A = Ane - (Ank+Anv+Anb+Ans)
    return A
    
def get_annuity_pv(area,pel_grid):
    obperiod = 10
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b
    deperiod = 25 #http://www.bosch-solarenergy.de/media/remote_media_se/alle_pdfs/broschueren/endverbraucher/solarstrom/Bosch_Solar_Energy_Solarstrom_en.pdf
    finst = 0.005
    fwins = 0.01
    effop = 5
    be = b
    a = get_annuity_factor(q,obperiod)
    #Capital related costs for the PV include price of purchase and installation costs
    CRC = 200*area + 500 + 65*area
    Ank = get_ank(CRC,r,q,obperiod,deperiod)

    #Demand related costs include price of fuel to generate the required electricity
    DRC = 0
    Anv = DRC*a*bv

    #Operation related costs include maintanance and repair
    ORC = 30*effop
    Ain = CRC*(finst+fwins)
    Anb = ORC*a*bb + Ain*a*bi
    #Other costs
    Ans = 0
    
    #Proceeds http://www.ise.fraunhofer.de/en/publications/veroeffentlichungen-pdf-dateien-en/studien-und-konzeptpapiere/recent-facts-about-photovoltaics-in-germany.pdf      
    #When battery is present, all of the electricity is assumed to be consumed in-house. 
    #Grid electricity being 35 cents, 35 cents are saved.
    #When there is no battery present, it is exported to the grid with FIT of 12 cents

    E1 = pel_grid*0.12
    
    Ane = E1*a*be
    A = Ane - (Ank+Anv+Anb+Ans)
    return A 

def get_annuity_battery(capacity):
    capacity = capacity*1000 #In Watts
    obperiod = 10
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b
    deperiod = 5 #Depreciation period is only 5 years
    finst = 0.005
    fwins = 0.01
    effop = 0
    
    a = get_annuity_factor(q,obperiod)
    #Capital related costs for the battery include price of purchase and installation costs
    CRC = 0.1308*capacity - 21.774
    Ank = get_ank(CRC,r,q,obperiod,deperiod)

    #Demand related costs include price of fuel to generate the required heat
    DRC = 0
    Anv = DRC*a*bv

    #Operation related costs include maintanance and repair
    ORC = 30*effop
    Ain = CRC*(finst+fwins)
    Anb = ORC*a*bb + Ain*a*bi
    #Other costs
    Ans = 0
    
    #Proceeds 
    Ane = 0
    
    A = Ane - (Ank+Anv+Anb+Ans)
    return A
    
def get_annuity_thstorage(capacity):
    obperiod = 10
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b
    deperiod = 15
    finst = 0.02
    fwins = 0.01
    effop = 0

    #converting kWh to liters 4180 is c of water and 40 is the temp difference
    capacity = capacity*3600000/(4180*40)
 
    a = get_annuity_factor(q,obperiod)
    
    if capacity>1000:
        bonus = 250*capacity/1000
    else:
        bonus = 0
    #Capital related costs for the storage include price of purchase and installation costs
    CRC = 1.0912*capacity + 367.92 -bonus
    Ank = get_ank(CRC,r,q,obperiod,deperiod)


    #Demand related costs include price of fuel to generate the required heat
    DRC = 0
    Anv = DRC*a*bv

    #Operation related costs include maintanance and repair
    ORC = 30*effop
    Ain = CRC*(finst+fwins)
    Anb = ORC*a*bb + Ain*a*bi
    
    #Other costs
    Ans = 0
    #Proceeds 
    Ane = 0
    print CRC,Ank,Anv,Anb,Ans,Ane,bonus
    A = Ane - (Ank+Anv+Anb+Ans)
    return A
    
def get_annuity_elheat(capacity,Qyearly):
    obperiod = 10
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b
    deperiod = 15
    finst = 0.01
    fwins = 0.01
    effop = 5
    
    #Changing kW to W
    capacity = capacity*1000
    
    a = get_annuity_factor(q,obperiod)
    #Capital related costs for the electrical heater include price of purchase and installation costs
    CRC = 53.938*capacity**0.2685
    Ank = get_ank(CRC,r,q,obperiod,deperiod)

    #Demand related costs include price of fuel to generate the required heat
    DRC = Qyearly*0.26
    Anv = DRC*a*bv

    #Operation related costs include maintanance and repair
    ORC = 30*effop
    Ain = CRC*(finst+fwins)
    Anb = ORC*a*bb + Ain*a*bi
    #Other costs
    Ans = 0
    
    #Proceeds 
    Ane = 0
#    print 'Electric Heater'
#    print 'A0=Investment amount=',CRC
#    print 'Ank=Capital Realted Annuity=',Ank
#    print 'Anv=Demand Related Annuity=',Anv
#    print 'Anb=Operation Related Annuity=',Anb
#    print 'Ane=Proceeds through feed-in-tariff=',Ane    
    A = Ane - (Ank+Anv+Anb+Ans)
    return A 
    
def get_annuity_CHP(capacity,Qyearly,storage):
    el_capacity = capacity*0.3
    obperiod = 10
    q = 1.07
    r = 1.03
    b = get_b(q,r,obperiod)
    bv = b
    bb = get_b(q,1.02,obperiod)
    bi = b
    deperiod = 15
    effop = 10
    fwins = 0.01
    finst = 0.01
    be = b
    
    #This is price per kWel.
    bonus = get_CHP_bonus(el_capacity)
    
    #Taken fron ASUE data of 2015. Gives euro/kwel
    if el_capacity <= 10:
        CRC = 9.585*el_capacity**(-0.542)
    elif el_capacity>10 and el_capacity <=100:
        CRC = 5.438*el_capacity**(-0.351)
    elif el_capacity>100 and el_capacity <=1000:
        CRC = 4.907*el_capacity**(-0.352)
    else:
        CRC = 1.7*460.89*el_capacity**(-0.015)
        
    #Multiplying to get euros
    CRC = CRC*el_capacity*1000*1.4
    a = get_annuity_factor(q,obperiod)

    #Capital related costs for the CHP include price of purchase and installation costs
    if storage:
        Ank = get_ank(CRC-bonus,r,q,obperiod,deperiod)
    else:
        Ank = get_ank(CRC,r,q,obperiod,deperiod)

    #Demand related costs include price of fuel to generate the required heat
    DRC = 0.0615*Qyearly/.6
    Anv = DRC*a*bv

    #Operation related costs include maintanance and repair
    #ORC = 30*effop
    #Ain = CRC*(finst+fwins)
    #Anb = ORC*a*bb + Ain*a*bi
    if el_capacity>0 and el_capacity<10:
        ORC=(3.2619*el_capacity**0.1866)*Qyearly/0.6*.3/100
    elif el_capacity>=10 and el_capacity<100:
        ORC=(6.6626*el_capacity**-0.25)*Qyearly/0.6*.3/100
    elif el_capacity>=100 and el_capacity<1000:
        ORC=(6.2728*el_capacity**-0.283)*Qyearly/0.6*.3/100
    Anb=ORC*a*bb
    #Other costs
    Ans = 0
    
    #Proceeds
    E1= Qyearly/0.6*0.3*0.10#*0.5 + 0.5*.3)
    Ane = E1*a*be
    E12 = Qyearly/0.6*0.3*.3
    Ane2 = E12*a*be
    A = Ane - (Ank+Anv+Anb+Ans)
#    print 'CHP'
#    print 'A0=Investment amount=',CRC
#    print 'Ank=Capital Realted Annuity=',Ank
#    print 'Anv=Demand Related Annuity=',Anv
#    print 'Anb=Operation Related Annuity=',Anb
#    print 'Ane=Proceeds through feed-in-tariff=',Ane
#    print 'Ane=Proceeds total self consumption=',Ane2    
    return a,CRC,bonus,Ank,Anv,Anb,Ane,A 


def get_CHP_bonus(capacity):
    if capacity>0 and capacity<=1:
        factor = 1900
    elif capacity>1 and capacity<=4:
        factor = 300
    elif capacity>4 and capacity<=10:
        factor = 100
    elif capacity>10 and capacity<=20:
        factor = 10
    else:
        factor = 1000000000
        
    if capacity>0 and capacity<=1:
        bonus = 1900
    elif capacity>1 and capacity<=2:
        bonus = 1900 + (capacity-1)*factor
    elif capacity>2 and capacity<=5:
        bonus = 2200 + (math.floor(capacity)-2)*300 + (capacity-math.floor(capacity))*factor
    elif capacity>5 and capacity<=11:
        bonus = 2900 + (math.floor(capacity)-5)*100 + (capacity-math.floor(capacity))*factor
    elif capacity>11 and capacity<=20:
        bonus = 3410 + (math.floor(capacity)-11)*10 + (capacity-math.floor(capacity))*factor
    else:
        bonus = 100000000000000000000
    
    bonus = bonus*1.25
    return bonus

print get_b(1.07,1.03,10)
print get_b(1.07,1.02,10)

#print get_annuity_CHP(8,15143,True)
#print get_annuity_thstorage(23.2222222222)
#print get_annuity_boiler(11,15437,True)
#print get_annuity_factor(1.07,10)#
#print get_annuity_elheat(0.3333,2920)
#print get_annuity_thstorage(database.get_thstorage_capacity(24))