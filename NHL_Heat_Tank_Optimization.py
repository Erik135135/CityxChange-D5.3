
from msilib import Binary

from pyomo.environ import *
import datetime as dt


from NHL_Results_v10 import write_results_nhl
from read_file_1min import read_file
from excel_csv import read_file1


def nhl():
    # Create csv files to be used in the model
    read_file("Data_periods_v1_1min_bid.xlsx", periods=24)
    read_file1('Data_heatpumps.xlsx', periods_on=False) # contains parameters with heat pump number as index
    read_file("TT.xlsx", periods_on=False)
    read_file("TI.xlsx", periods_on=False)
    read_file("RF506.xlsx", periods_on=False)
    read_file("MaxEl.xlsx", periods_on=False)
    read_file("MaxHeat.xlsx", periods_on=False)
    read_file("RF401.xlsx", periods_on=False)



    index = 1
    ########################################## Optimization model ##########################################################
    model = AbstractModel()

    model.periods = Set()
    model.heatpumps = Set()

    # Parameters
    model.FixDemEl = Param(model.periods, mutable=True) # all el equipments that is not flexible in hour model
    model.CEl = Param(model.periods, mutable=True)
    model.CostHeat = Param(model.periods, mutable=True)
    model.Date = Param(model.periods)
    model.MinTemp = Param(model.periods, mutable=True)
    model.TOut = Param(model.periods, mutable=True)

    model.cp = Param(default=1.16) # heat capacity for water [Wh/kgC]
    model.massflow = Param(default=4) # mass flow [l/s], I think this should be change to [l/h] (pressure difference (RP501-RD401)


    # Temperatures


    #model.Thp = Param(model.periods) # Temperature transferred to system from heat pumps (Parameter?)
    #model.Ttank = Param(model.periods) # Parameter???

    model.Tin_start = Param()

    model.RT401_start = Param()
    #model.RT501_start = Param(default=50)
    model.RT506_start = Param()
    model.CollectHeat = Param()
    model.CollectEl = Param()

    # Heat Pumps
    model.COP = Param(model.heatpumps, default=0)
    model.HPCap = Param(model.heatpumps, default=0)

    # Variables(El)
    model.DEl = Var(model.periods, within=NonNegativeReals)
    model.maxEl = Var(within=NonNegativeReals)
        #Heat pumps
    model.HPEl = Var(model.periods, model.heatpumps, within=NonNegativeReals)

    # Variables (Heat)
    model.tempLoss = Var(model.periods, within=NonNegativeReals)
    model.DHeat = Var(model.periods, within=NonNegativeReals)
    model.maxHeat = Var(within=NonNegativeReals)
    model.DRQImport = Var(model.periods, within=NonNegativeReals)
    model.DTQImport = Var(model.periods, within=NonNegativeReals)
        # System
    model.RT401 = Var(model.periods, within=NonNegativeReals)
    model.RT506 = Var(model.periods, within=NonNegativeReals)
    model.RT501 = Var(model.periods, within=NonNegativeReals)
    model.DRQNeeded = Var(model.periods, within=NonNegativeReals)
    model.systemLoss = Var(model.periods, within=NonNegativeReals)
    model.Ttank = Var(model.periods, within=NonNegativeReals)
        #Building
    model.RTTbuilding = Var(model.periods, within=NonNegativeReals)
    model.Tbinside = Var(model.periods, within=NonNegativeReals)
    model.temp_loss = Var(model.periods, within=NonNegativeReals)
    model.Tin = Var(model.periods, within=NonNegativeReals)
    model.Tin_increase = Var(model.periods, within=NonNegativeReals)



        #HeatPumps
    model.HRQneeded = Var(model.periods, model.heatpumps, within=NonNegativeReals)
    model.HRQimported = Var(model.periods, model.heatpumps, within=NonNegativeReals)
    model.Thp = Var(model.periods, model.heatpumps, within=NonNegativeReals)
        # Tank

    model.tempSoCTank_start = Param()
    model.VolumTank = Param(default=10000) #liter
    model.Spes_heat = Param(default=4.16)
    model.RT502_start = Param(default=47)

    model.socTankQ = Var(model.periods)
    model.chargeDTQ = Var(model.periods, within=NonNegativeReals)
    model.dischargeTRQ = Var(model.periods, within=NonNegativeReals)
    model.dischargeNeededTRQ = Var(model.periods, within=NonNegativeReals)
    model.tankSoCQ_start = Var(within=NonNegativeReals)
    model.chargeNeededDTQ = Var(model.periods, within=NonNegativeReals)
    model.dischargeImportTRQ = Var(model.periods, within=NonNegativeReals)
    model.DTQimported = Var(model.periods, within=NonNegativeReals)

    # Binary Variables
    model.b1 = Var(within=Binary)
    model.b2 = Var(within=Binary)
    model.b3 = Var(within=Binary)
    model.b4 = Var(within=Binary)
    model.b5 = Var(within=Binary)
    model.b6 = Var(within=Binary)
    model.b7 = Var(within=Binary)
    model.b8 = Var(within=Binary)


    model.SB501 = Var(model.periods, within=Binary)
    model.SB502 = Var(model.periods, within=Binary)
    model.SB503 = Var(model.periods, within=Binary)

    model.binary1 = Var(model.periods, within=Binary)

    # Constants
    model.HP1Max = Param(default=46.4)  # Capacity of HP1 [kW]
    model.HP2Max = Param(default=26.5)  # Capacity of HP1 [kW]


    ################################## Objective function #################################################################

    def obj_function(model):
        return sum(model.DEl[t] * model.CEl[t] for t in model.periods) + (sum((model.DHeat[t]) * model.CostHeat[t] for t in model.periods)) + \
               ((45 * model.b1 + 40*model.b2 + 35*model.b3 + 30 *model.b4)) * model.maxHeat + \
               ((45 * model.b5 + 42*model.b6 + 38*model.b7 + 35*model.b8)) * model.maxEl + sum(30000*model.binary1[t] for t in model.periods)
    model.obj_function = Objective(rule=obj_function, sense=minimize)

    # Heat demand and power cost
    def total_demand_heat (model,t):
        return model.DHeat[t] == model.DRQImport[t] + model.DTQimported[t]
    model.total_demand_heat = Constraint(model.periods, rule = total_demand_heat)

    def maxheat (model, t): # Modify acording to Geologi
        return model.DHeat[t] <= model.maxHeat
    model.MaxHeat = Constraint(model.periods, rule=maxheat)

    def maxheat1 (model): # Modify acording to Geologi
        return model.CollectHeat <= model.maxHeat
    model.MaxHeat1 = Constraint(rule=maxheat1)


    def costHeatbin(model):
        return model.b1 + model.b2 + model.b3 + model.b4 == 1
    model.CostHeatbin = Constraint(rule=costHeatbin)

    def heat_import(model): #Change numbers to correct max heat (Different summer and winter months)
        return (model.maxHeat-0)*model.b1 + (model.maxHeat - 200) *model.b2 + (model.maxHeat - 500)*model.b3 +  (model.maxHeat - 800)*model.b4 >= 0
    model.Heat_importer = Constraint(rule=heat_import)

    #Electricity demand and power cost

    model.elhp = Var(model.periods, model.heatpumps, within=NonNegativeReals)

    def heat_pump(model, t):
        if value(model.TOut[t]) >= -8:
            return model.Thp[t, 1] <= (model.HPCap[1] *model.COP[1])/60
        else:
            return model.Thp[t, 1] == 0

    model.heatPump_act = Constraint(model.periods, rule=heat_pump)

    def heat_pump1(model, t):
        if value(model.TOut[t]) >= -8:
            return model.Thp[t, 2] <= model.HPCap[2] * model.COP[2]/60
        else:
            return model.Thp[t, 2] == 0
    model.heatPump_act1 = Constraint(model.periods, rule=heat_pump1)

    def demand_el(model, t):
        return model.DEl[t] == model.FixDemEl[t] + (model.Thp[t, 1])*60/model.COP[1] + (model.Thp[t, 2])*60/model.COP[2]
    model.demand_el = Constraint(model.periods, rule=demand_el)

    def maxEl(model, t): #Geologi version
        return model.DEl[t] <= model.maxEl
    model.MaxElec = Constraint(model.periods, rule=maxEl)

    def maxEl1(model, t): #Geologi version
        return model.CollectEl <= model.maxEl
    model.MaxElec1 = Constraint(model.periods, rule=maxEl1)

    def CElbin(model):
        return model.b5 + model.b6 + model.b7 + model.b8 == 1
    model.CElbin = Constraint(rule=CElbin)

    def el_import(model):
        return (model.maxEl-0)*model.b5 + (model.maxEl - 200) *model.b6 + (model.maxEl - 500)*model.b7 + (model.maxEl - 800)*model.b8 >= 0
    model.El_importer = Constraint(rule=el_import)

    #################################################### Energy balance ###################################################


    def rt401(model,t):
        if value(model.TOut[t])>= 15:
            return model.RT401[t] + 5*model.binary1[t] >= 30
        if 10 <= value(model.TOut[t])<= 15:
            return model.RT401[t] + 5*model.binary1[t] >= 40 - (model.TOut[t] - 10)*2
        elif -8 <= value(model.TOut[t]) < 10:
            return model.RT401[t] + 5*model.binary1[t] >= 50 - (model.TOut[t])
        else:
            return model.RT401[t] + 5 * model.binary1[t] >= 65
    model.rt401 = Constraint(model.periods, rule=rt401)

    def heat_DR_from_temperature(model,t):
            if t == index:
                return model.DRQNeeded[t] == (model.RT401[t] - model.RT506_start)*60
            else:
                return model.DRQNeeded[t] == (model.RT401[t] - model.RT506[t-1])*60
    model.heat_DR_from_temperature = Constraint(model.periods, rule=heat_DR_from_temperature)


    def temperature_to_heat(model,t):
        return (model.DRQImport[t]) == (model.DRQNeeded[t]/0.95) # preliminary conversion rate = epsilon 1
    model.temperature_to_heat = Constraint(model.periods, rule=temperature_to_heat)

    def temperature_to_heat_max(model,t):
        return (model.DRQImport[t]) <= 750
    model.TempMaxImport = Constraint(model.periods, rule=temperature_to_heat_max)


    def rt506(model, t):
        if t == index:
            return model.RT506[t] == model.RT506_start + model.DRQNeeded[t]/60 - model.RTTbuilding[t] + (model.Ttank[t]) + sum((model.Thp[t, i]) for i in model.heatpumps) - model.systemLoss[t]
        elif 1 < t < 25:
            return model.RT506[t] == model.RT506[t-1] + model.DRQNeeded[t]/60 - model.RTTbuilding[t] + (model.Ttank[t]) + sum((model.Thp[t, i]) for i in model.heatpumps) - model.systemLoss[t]
        else:
            return model.RT506[t] == model.RT506[t-1]

    model.rt506 = Constraint(model.periods, rule=rt506)


    # Loss in the system

    def loss_heat_system(model, t):
        return model.systemLoss[t] == 0.03 * model.RT401[t]
    model.returnheatSystem = Constraint(model.periods, rule=loss_heat_system)


    # Building
    def temperature_transfer_building(model, t):
        if t == index:
            return model.Tin[t] == model.Tin_start + model.Tin_increase[t] - model.temp_loss[t]
        else:
            return model.Tin[t] == model.Tin[t - 1] + model.Tin_increase[t] - model.temp_loss[t]
    model.temperature_transfer_building = Constraint(model.periods, rule=temperature_transfer_building)


    def temperature_transfer(model,t):
        return model.Tin_increase[t] == model.RTTbuilding[t]*0.34
    model.TemperatureTransfer = Constraint(model.periods, rule=temperature_transfer)

#        return model.Tin_increase[t] == model.RTTbuilding[t]*0.317155214

    def temperature_min(model,t):
        return model.Tin[t] + 2*model.binary1[t] >= model.MinTemp[t]
    model.temperature_min = Constraint(model.periods, rule=temperature_min)

    def temperature_loss(model,t):
        return model.temp_loss[t] == 0.1*(model.Tin[t]-model.TOut[t]) # Change to a function of TempInside and TOut
    model.temperature_loss = Constraint(model.periods, rule=temperature_loss)


    def soc_temp_to_energi(model):
        return model.tankSoCQ_start == model.tempSoCTank_start*model.VolumTank*model.Spes_heat*0.000277777778

    model.SoCTempToEnergy = Constraint(rule=soc_temp_to_energi)

    model.TempTank = Var(model.periods, within=NonNegativeReals)
    model.TempTankCharge = Var(model.periods, within=NonNegativeReals)
    model.TempTankDischarge = Var(model.periods, within=NonNegativeReals)
    model.ChDischEff = Param(default=0.95)
    model.LossTempTankPer = Param(default=0.001)
    model.MaxTempTank = Param(default=90)
    model.MinTempTank = Param(default=40)

    def state_charge_tank_energi(model,t):
        if t == index:
            return model.TempTank[t] == model.tempSoCTank_start + model.TempTankCharge[t]*model.ChDischEff  - model.TempTankDischarge[t]/model.ChDischEff - model.LossTempTankPer*model.tempSoCTank_start
        else:
            return model.TempTank[t] == model.TempTank[t-1] + model.TempTankCharge[t]*model.ChDischEff  - model.TempTankDischarge[t]/model.ChDischEff  - model.LossTempTankPer*model.TempTank[t]

    model.stateOfChargeEnergi = Constraint(model.periods, rule=state_charge_tank_energi)

    def tempTank(model,t):
        return model.socTankQ[t] == ((model.TempTank[t]) - model.RT506[t])*model.VolumTank*model.Spes_heat*0.000277777778
    model.TempTANK_1 = Constraint(model.periods, rule=tempTank)


    def maxtempTank(model,t):
        return model.TempTank[t] - 5*model.binary1[t] <= MaxTempTank
    model.MaxTempTank = Constraint(model.periods, rule=maxtempTank)

    def mintempTank(model,t):
        return model.TempTank[t] + 5*model.binary1[t] >= model.MinTempTank
    model.MinTempTank = Constraint(model.periods, rule=mintempTank)

    def heat_imported_system_tank(model,t):
            return model.DTQimported[t] == (model.TempTankCharge[t])*model.VolumTank*model.Spes_heat*0.000277777778  #epsilon 4
    model.heat_imported_system_tank = Constraint(model.periods, rule=heat_imported_system_tank)

    def heat_imported_system_tank1(model,t):
        return model.TempTankCharge[t] <= 7   #epsilon 4
    model.heat_imported_system_tank1 = Constraint(model.periods, rule=heat_imported_system_tank1)

    def heat_imported_tank(model,t):
        return model.dischargeImportTRQ[t] == ((model.TempTankDischarge[t])*model.VolumTank*model.Spes_heat*0.000277777778)/60
    model.heat_imported_tank = Constraint(model.periods, rule=heat_imported_tank)

    def heat_export_system_tank1(model,t):
        return model.TempTankDischarge[t] <= (model.TempTank[t] - model.RT506[t])*0.3*model.SB502[t]
    model.heat_imported_system_tank12 = Constraint(model.periods, rule=heat_export_system_tank1)

    def energy_needed_tank1(model,t):
        return model.Ttank[t] == (model.dischargeImportTRQ[t])
    model.EnergyNeededTank1 = Constraint(model.periods, rule=energy_needed_tank1)



    ################################################### Load Data #########################################################
    data = DataPortal()
    data.load(filename='Data/Data_periods.csv', select=("Index", 'Time', 'To', 'PriceEl', 'PriceH', 'fd_elec', 'min_temp'), param=(model.Date, model.TOut, model.CEl, model.CostHeat, model.FixDemEl, model.MinTemp), index=model.periods)
    data.load(filename='Data/Data_heatpumps.csv', select=('Pumpe', 'COP', 'CapHP'), param=(model.COP, model.HPCap), index=model.heatpumps)
    data.load(filename=r'Data\TT.csv', param=model.tempSoCTank_start)
    data.load(filename=r'Data\TI.csv', param=model.Tin_start)
    data.load(filename=r'Data\RF506.csv', param=model.RT506_start)
    data.load(filename=r'Data\MaxEl.csv', param=model.CollectEl)
    data.load(filename=r'Data\MaxHeat.csv', param=model.CollectHeat)
    data.load(filename=r'Data\RF401.csv', param=model.RT401_start)



    ######################################################### Solver ######################################################
    instance = model.create_instance(data)  # load parameters
    #results = SolverFactory("gurobi", Verbose=True).solve(instance, tee=True)
    solver = SolverFactory("gurobi")
    #solver.options['mipgap'] = 0.01
    solver.options['TimeLimit'] = 160 # Solving a model instance
    results = solver.solve(instance, tee=True)
    instance.solutions.load_from(results)  # Loading solution into instance
    write_results_nhl(instance)

    return instance

nhl()