
from msilib import Binary

from pyomo.environ import *
from pyomo.repn.plugins.baron_writer import NonNegativeReals, NonNegativeIntegers


def geologi(value1, value2, yourdate, now1, value5):
    from Geologi_Results import write_results_geo
    from read_file_v1 import read_file_v1
    from Test_log_data import Code_write
    from forecasts_update import forecast_and_data_update
    import datetime as dt

    # Read excel files created from forecast and previous data
    read_file_v1("Parameters.xlsx", periods_on=False)
    read_file_v1("Data_periods_v1.xlsx", periods=120)
    read_file_v1("CollectEl.xlsx", periods_on=False)
    read_file_v1("HeatEx.xlsx", periods_on=False)

    index=1

    model = AbstractModel()

    #### We have to add the fan System losses that occurs through the heat exchanger.

    #Sets


    model.Periods = Set() # Set of time steps

    model.TimeEff = RangeSet(1,85) # Set of time steps for hourly constraints


    #Parameters
        # Forecast data
    model.TempOutside = Param(model.Periods, mutable=True) #Forecast for the temperature outside [°C]
    model.FixedDemand = Param(model.Periods, mutable=True) #Includes de demand for appliances, fan, light and heat panels [kWh/12h]
    model.Cost = Param(model.Periods, mutable=True) #Cost of the electricity for each period [NOK/kWh]
    model.Date = Param(model.Periods) # Time and date of the predictions [01.01.0000 00:00:00]

    # Live data
    model.RealTemperature = Param(initialize=value1) # Real temperature inside Geologi (Live data)
    model.HeatExhangerPer = Param(initialize=value2) # Real usage of  heat exchanger (Live data)
    model.RealTemperatureInto = Param(initialize=value5) # Real temperature of air into Geologi from ventilation system (Live data)
    model.RealTemperatureOutside = Param() # Real temperature of outside around NTNU Gløshaugen (Live data)

    #Constants

    model.fanPower = Param(default=5) # Power of the fan in the ventilation system (constant air flow) [kW]
    model.CollectEl = Param(initialize=0) # Collected maximum peak of electricity from previous timesteps
    model.HeatEx = Param(initialize=0) #
    model.Battery_capacity = Param() #The size of the battery [kW]


    #Variables
    model.battery = Var(model.Periods, within=NonNegativeReals)
    model.totalDemand = Var(model.Periods, within=NonNegativeReals)
    model.tempInside = Var(model.Periods, within=NonNegativeReals)
    model.tempLoss = Var(model.Periods, within=NonNegativeReals) #We assume there is no loss at the starting time, this can change depending on the dataset
    model.tempBattery = Var(model.Periods, within=NonNegativeReals)
    model.heatexch_on = Var(model.Periods, within=Binary)
    model.PercentageSetPoint =Var(model.Periods, within=NonNegativeIntegers)
    model.fan = Var(model.Periods, within=NonNegativeReals)
    model.tempSetPoint = Var(model.Periods, within=NonNegativeReals)
    model.battery_percentage = Var(model.Periods, within=NonNegativeReals)
    model.el = Var(model.TimeEff, within=NonNegativeReals)
    model.tempinto = Var(model.Periods, within=NonNegativeReals)
    model.aboveTemp = Var(model.Periods, within=NonNegativeReals)
    model.maxEl = Var(within=NonNegativeReals)
    model.TemperatureInside = Var(within=NonNegativeReals)


    model.b1 = Var(within=Binary)
    model.b2 = Var(within=Binary)
    model.b3 = Var(within=Binary)
    model.maxEl = Var(within=NonNegativeReals)
    model.binary = Var(model.Periods, within=Binary)
    model.binaryPunish = Var(model.Periods, within=Binary)



    #JB01  = Difference between TF01 + (TF04 and TF 05 /TF04 - TF01)

    #OBJECTIVE FUNCTION
    def obj_function(model):
        return sum(model.totalDemand[t] * (model.Cost[t]/1000) + (model.binaryPunish[t]*500)
                   for t in model.Periods) + ((45*model.b1 + 42*model.b2 + 38*model.b3)/30) * model.maxEl
    model.obj_function = Objective(rule=obj_function, sense=minimize)

#+ sum(50*model.binary_temp[t] for t in model.Periods)

    #CONSTRAINTS
    def tot_demand(model, time):
        return model.totalDemand[time] == model.FixedDemand[time] + model.battery[time] + model.fan[time]
    model.tot_demand = Constraint(model.Periods, rule=tot_demand)


    # Electricity peaks
    def max_el2(model,n):
        return ((model.totalDemand[n] + model.totalDemand[n+1] + model.totalDemand[n+2] +
                   model.totalDemand[n+3] + model.totalDemand[n+4] + model.totalDemand[n+5] + model.totalDemand[n+6] +
                   model.totalDemand[n+7] + model.totalDemand[n+8] + model.totalDemand[n+9] + model.totalDemand[n+10] +
                   model.totalDemand[n+11]))/12 == model.el[n]

    model.MaxEl2_1 = Constraint(model.TimeEff, rule=max_el2)


    def max_el(model, n):
        return model.el[n] <= model.maxEl
    model.MaxEl = Constraint(model.TimeEff, rule=max_el)

    def max_el2(model):
        return model.CollectEl <= model.maxEl
    model.MaxEl12 = Constraint(rule=max_el2)

    def effect_cost(model):
        return model.b1 + model.b2 + model.b3 == 1
    model.EffectCost = Constraint(model.Periods, rule=effect_cost)

    def heat_import(model):
        return (model.maxEl- 0) * model.b1 + (model.maxEl - 100) * model.b2 + (model.maxEl - 400) * model.b3 >= 0
    model.Heat_importer = Constraint(rule=heat_import)

    # Battery
    def per_to_kwh(model,time):
        return model.battery[time] == 1.15*model.battery_percentage[time]
    model.PerToKwh = Constraint(model.Periods, rule=per_to_kwh)

    def cap_battery(model, time):
        return model.battery_percentage[time] <= 100
    model.cap_battery = Constraint(model.Periods, rule=cap_battery)

    # fan
    def low_cap_battery(model, time):
        return model.fan[time] == model.fanPower * model.heatexch_on[time]
    model.LowCapBattery = Constraint(model.Periods, rule=low_cap_battery)

    def vent_on(model, time):
        if 0 <= dt.date.today().weekday() <= 3:
            if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("22:00", "%H:%M").time()):
                return model.heatexch_on[time] == 1
            else:
                return model.heatexch_on[time] == 0
        else:
            if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                return model.heatexch_on[time] == 1
            else:
                return model.heatexch_on[time] == 0

    model.vent_on = Constraint(model.Periods, rule=vent_on)

    # Building


    def sensitivity_Temp(model):
        if value(model.TempOutside[1]) <= -5:
            if 0 <= dt.date.today().weekday() <= 4:
                if (20.6 - 0.2) <= value(model.RealTemperature - 0.5 - model.TempOutside[1]*0.1) <= (20.6 + 0.2):
                    return model.TemperatureInside == 20.6
                else:
                    return model.TemperatureInside == model.RealTemperature - 0.5 - model.TempOutside[1]*0.1
            else:
                if (20.4 - 0.2) <= value(model.RealTemperature - 0.5 - model.TempOutside[1]*0.1) <= (20.4 + 0.2):
                    return model.TemperatureInside == 20.4
                else:
                    return model.TemperatureInside == model.RealTemperature - 0.5 - model.TempOutside[1] * 0.1
        elif 0 >= value(model.TempOutside[1]) > -5:
            if 0 <= dt.date.today().weekday() <= 4:
                if (20.7 - 0.2) <= value(model.RealTemperature - 0.3 - model.TempOutside[1]*0.06) <= (20.7 + 0.2):
                    return model.TemperatureInside == 20.7
                else:
                    return model.TemperatureInside == model.RealTemperature - 0.3 - model.TempOutside[1]*0.06
            else:
                if (20.5 - 0.2) <= value(model.RealTemperature - 0.3 - model.TempOutside[1]*0.06) <= (20.5 + 0.2):
                    return model.TemperatureInside == 20.5
                else:
                    return model.TemperatureInside == model.RealTemperature - 0.3 - model.TempOutside[1] * 0.06
        else:
            if 0 <= dt.date.today().weekday() <= 4:
                if (21 - 0.2) <= value(model.RealTemperature - 0.4) <= (21 + 0.2):
                    return model.TemperatureInside == 20.8
                else:
                    return model.TemperatureInside == model.RealTemperature - 0.6
            else:
                if (20.6 - 0.2) <= value(model.RealTemperature - 0.2) <= (20.6 + 0.2):
                    return model.TemperatureInside == 20.6
                else:
                    return model.TemperatureInside == model.RealTemperature - 0.4
    model.SenSitivity_Temp = Constraint(rule=sensitivity_Temp)



    def temp_inside(model, time):
        if 0 <= dt.date.today().weekday() <= 3:
            if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("22:00", "%H:%M").time()):
                if time == index:
                    return model.tempInside[time] == (model.TemperatureInside) + 0.15 * (model.tempinto[time] - (model.TemperatureInside)) - ((model.TemperatureInside) - model.TempOutside[time]) * 0.00050
                else:
                    return model.tempInside[time] == model.tempInside[time - 1] + 0.15 * (model.tempinto[time] - (model.tempInside[time - 1])) - ((model.tempInside[time - 1]) - model.TempOutside[time]) * 0.00050
            else:
                if time == index:
                    return model.tempInside[time] == (model.TemperatureInside) - ((model.TemperatureInside) - model.TempOutside[time]) * 0.00050
                else:
                    return model.tempInside[time] == model.tempInside[time - 1] - ((model.tempInside[time - 1]) - model.TempOutside[time]) * 0.00050
        else:
            if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                if time == index:
                    return model.tempInside[time] == (model.TemperatureInside) + 0.15 * (model.tempinto[time] - (model.TemperatureInside)) - ((model.TemperatureInside) - model.TempOutside[time]) * 0.00050
                else:
                    return model.tempInside[time] == model.tempInside[time - 1] + 0.15 * (model.tempinto[time] - (model.tempInside[time - 1])) - ((model.tempInside[time - 1]) - model.TempOutside[time]) * 0.00050
            else:
                if time == index:
                    return model.tempInside[time] == (model.TemperatureInside) - ((model.TemperatureInside) - model.TempOutside[time]) * 0.0005
                else:
                    return model.tempInside[time] == model.tempInside[time - 1] - ((model.tempInside[time - 1]) - model.TempOutside[time]) * 0.00050

    model.loss_inside = Constraint(model.Periods, rule=temp_inside)



    def min_temp_inside(model, time):
        if value(model.TempOutside[time]) <= -5:
            if 0 <= dt.date.today().weekday() <= 3:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00","%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("20:00", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.6
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.6
                elif (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("20:05","%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("22:00","%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.3
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.3
                else:
                    return 26 >= model.tempInside[time] >= 13
            elif dt.date.today().weekday() == 4:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.6
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.6
                else:
                    return 26 >= model.tempInside[time] >= 13
            else:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.3
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.3
                else:
                    return 26 >= model.tempInside[time] >= 13
        elif 0 >= value(model.TempOutside[1]) > -5:
            if 0 <= dt.date.today().weekday() <= 3:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("20:00", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.7
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.7
                elif (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("20:05", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("22:00", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.4
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.4
                else:
                    return 26 >= model.tempInside[time] >= 13
            elif dt.date.today().weekday() == 4:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.7
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.7
                else:
                    return 26 >= model.tempInside[time] >= 13
            else:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.5
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.5
                else:
                    return 26 >= model.tempInside[time] >= 13
        else:
            if 0 <= dt.date.today().weekday() <= 3:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00","%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("20:00", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.8
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.8
                elif (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("20:05", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("22:00", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.6
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.6
                else:
                    return 26 >= model.tempInside[time] >= 13
            elif dt.date.today().weekday() == 4:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.8
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.8
                else:
                    return 26 >= model.tempInside[time] >= 13
            else:
                if (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00", "%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]), "%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30", "%H:%M").time()):
                    if time == index:
                        return 26 >= model.TemperatureInside + 4 * model.binaryPunish[time] >= 20.6
                    else:
                        return 26 >= model.tempInside[time] + 4 * model.binaryPunish[time] >= 20.6
                else:
                    return 26 >= model.tempInside[time] >= 13
    model.min_temp_inside = Constraint(model.Periods, rule=min_temp_inside)


    def set_point(model, time):
        if value2 >= 98 or value(model.HeatEx) >= 98:
                if 0 <= dt.date.today().weekday() <= 3:
                    if (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00","%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("22:00","%H:%M").time()):
                        return 20 <= model.tempSetPoint[time] <= 22
                    else:
                        return model.tempSetPoint[time] == 20
                else:
                    if (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() > dt.datetime.strptime("07:00","%H:%M").time()) and (dt.datetime.strptime(value(model.Date[time]),"%Y-%m-%d %H:%M:%S").time() < dt.datetime.strptime("18:30","%H:%M").time()):
                        return 20 <= model.tempSetPoint[time] <= 22
                    else:
                        return model.tempSetPoint[time] == 20
        else:
            return model.tempSetPoint[time] == (model.heatexch_on[time]*21) + (1-model.heatexch_on[time])*20

    model.Set_point_ = Constraint(model.Periods, rule=set_point)


    def defin(model,time):
        if time == index:
            if value(model.TempOutside[time]) >= 21:
                return model.aboveTemp[time] == model.TempOutside[time] - model.RealTemperature
            else:
                return model.aboveTemp[time] == 0
        else:
            if value(model.TempOutside[time]) >= 21:
                return model.aboveTemp[time] == model.TempOutside[time] - model.tempInside[time]
            else:
                return model.aboveTemp[time] == 0

    model.defIn1 = Constraint(model.Periods, rule=defin)



    def setpoimt(model, time):
        if value2 >= 98 or value(model.HeatEx) >= 98:
            if time == index:
                return model.tempinto[time] == round(value(model.RealTemperatureInto)) + 0.2*(model.tempSetPoint[time] - round(value(model.RealTemperatureInto)))
            else:
                return model.tempinto[time] == model.tempinto[time-1] + 0.2 * (model.tempSetPoint[time] - model.tempinto[time-1])
        else:
            if value(model.TempOutside[time]) <= 9.3957:
                if time == index:
                    return model.tempinto[time] == round(value(model.RealTemperatureInto)) + 0.2*(model.tempSetPoint[time] - round(value(model.RealTemperatureInto)))
                else:
                    return model.tempinto[time] == model.tempinto[time-1] + 0.2 * (model.tempSetPoint[time] - model.tempinto[time-1])
            else:
                if time == index:
                    return model.tempinto[time] <= round(value(model.RealTemperatureInto)) + model.aboveTemp[time] + model.battery_percentage[time]/10
                else:
                    return model.tempinto[time] <= model.tempinto[time -1] + model.aboveTemp[time] + model.battery_percentage[time]/10

    model.SetPoint1 = Constraint(model.Periods, rule=setpoimt)


    def bat_set_point(model, time):
        if value2 >= 98 or value(model.HeatEx) >= 98:
            if value(model.TempOutside[time]) <= 9.3957:
                return model.tempSetPoint[time] <= ((9.3957 + model.battery_percentage[time]) / 0.8917 + (model.TempOutside[time])) * model.heatexch_on[time] + (1 - model.heatexch_on[time]) * 20
            else:
                return model.tempSetPoint[time] == (model.heatexch_on[time]*21) + (1-model.heatexch_on[time])*20
        else:
            return model.tempSetPoint[time] == (model.heatexch_on[time]*21) + (1-model.heatexch_on[time])*20
    model.SetPoint = Constraint(model.Periods, rule=bat_set_point)


    #DATA
    data = DataPortal()
    data.load(filename=r'Data\BatteryCapacity.csv', param=model.Battery_capacity)
    data.load(filename=r'Data\CollectEl.csv', param=model.CollectEl)
    data.load(filename=r'Data\Data_periods.csv', select=("Index","Time","To","Price","kW"), param=(model.Date, model.TempOutside, model.Cost, model.FixedDemand), index=model.Periods)
    data.load(filename=r'Data\HeatEx.csv', param=model.HeatEx)


    #SOLVER
    instance = model.create_instance(data) # load parameters
    results = SolverFactory('gurobi', Verbose=True).solve(instance, tee=True)
    instance.solutions.load_from(results)
    Code_write(instance)
    write_results_geo(instance, sheet_name="Results")

    return instance
