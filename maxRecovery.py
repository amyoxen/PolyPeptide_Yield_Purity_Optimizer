"""
The Simplified Crosyntropin Model Python Formulation for the PuLP Modeller

"""
# Import PuLP modeler functions
from pulp import *
import xlrd
import sys
import os

script_dir = os.path.dirname(__file__)
rel_path = 'data.xlsm'
abs_file_path = os.path.join(script_dir, rel_path)

workbook = xlrd.open_workbook(abs_file_path,'r')      
worksheet = workbook.sheet_by_index(1)

# Creates a list of the Fractions
Fractions = []
Concentrations = {}
i = 1
while worksheet.cell(i, 3).value != xlrd.empty_cell.value:
    fraction = worksheet.cell(i, 3).value
    concentration = worksheet.cell(i, 10).value
    Fractions.append (str(fraction))
    Concentrations[str(fraction)] = concentration
    i +=1

# Create the 'prob' variable to contain the problem data
prob = LpProblem("The Diet Problem", LpMaximize)

# A dictionary called 'fraction_vars' is created to contain the referenced Variables
fraction_vars = LpVariable.dicts("Ingr", Fractions, 0)

# The objective function is added to 'prob' first
prob += lpSum([Concentrations[i]*fraction_vars[i] for i in Fractions]), "Total Weight of Fractions"

for j in range(4,9):
    cd = {}
    i = 1
    while worksheet.cell(i, 3).value != xlrd.empty_cell.value:
        fraction = worksheet.cell(i, 3).value
        if worksheet.cell(i, j).value == xlrd.empty_cell.value: 
            cd[str(fraction)] = 0
        else:
            cd[str(fraction)] = worksheet.cell(i, j).value
        
        i +=1
   
    prob += lpSum([cd[i] * fraction_vars[i] for i in Fractions]) >= lpSum([fraction_vars[i] for i in Fractions])* float(worksheet.cell(45, j).value), "lower limit" + str(j)
    prob += lpSum([cd[i] * fraction_vars[i] for i in Fractions]) <= lpSum([fraction_vars[i] for i in Fractions])* float(worksheet.cell(46, j).value), "upper limit" + str(j)


weightLimit = {}
i = 1
while worksheet.cell(i, 3).value != xlrd.empty_cell.value:
    fraction = worksheet.cell(i, 3).value
    if worksheet.cell(i, 13).value == xlrd.empty_cell.value: 
        weightLimit[str(fraction)] = 0
    else:
        weightLimit[str(fraction)] = worksheet.cell(i, 13).value
    
    i  +=1

for i in Fractions:
    prob += fraction_vars[i] <= weightLimit[i], "WeightLimit" + str(i)

#prob.writeLP("CrosytropinModel.lp")

prob.solve()

# The status of the solution is printed to the screen
#print "Status:", LpStatus[prob.status]


#for v in prob.variables():
#    print v.name, "=", v.varValue
    
# The optimised objective function value is printed to the screen
#print "Total SumWeight of Fractions per meal = ", value(prob.objective)
#print "Total Fractions Servings per meal = ", sum([v.varValue for v in prob.variables()])

for v in prob.variables():
    sys.stdout.write (str(v.varValue))
    sys.stdout.write ("\n")