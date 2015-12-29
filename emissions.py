def get_emissions_boilers(Qyearly):
    emissions = 201.6*Qyearly
    return (emissions/1000)

def get_emissions_pv(pel_grid):
    emissions = -705*pel_grid
    return (emissions/1000)
    
def get_emissions_elheat(Qyearly):
    emissions = 595*Qyearly
    return (emissions/1000)

def get_emissions_CHP(Qyearly):
    #emissions = (180*Qyearly - 595*Qyearly*.3)/.6
    emissions = 2.5*Qyearly
    return (emissions/1000)
    