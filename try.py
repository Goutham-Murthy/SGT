import annuity

print annuity.get_annuity_CHP(2.58,8021)
print ('\n\n')
print annuity.get_annuity_CHP(2.58,9646)
print ('\n\n')
print annuity.get_annuity_boiler(11,8666)
print ('\n\n')
print annuity.get_annuity_boiler(11,7054)
print ('\n\n')

print annuity.get_annuity_CHP(2.58,12021)+annuity.get_annuity_boiler(11,8666)
print annuity.get_annuity_CHP(2.58,13646)+annuity.get_annuity_boiler(11,7054)