def get_pef(CHP=0,boiler=0,solth=0,elheat=0):
    pef = (CHP*0.7 + boiler*1.1 + solth*1 + elheat*2.8 )/(CHP+boiler+solth+elheat)
    return pef