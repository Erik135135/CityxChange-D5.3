
import pandas as pd
import random
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as cm
import matplotlib.patches as mpatches
import matplotlib.lines as mlines
import matplotlib.ticker as mtick
import matplotlib.pylab as pylab
from pyomo.environ import *


def get_periods(instance):
    Periods = instance.Periods
    for key in Periods:
        periods = int(Periods[key]) + 1
    return periods

def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False,
                       **to_excel_kwargs):
    """
    Append a DataFrame [df] to existing Excel file [filename]
    into [sheet_name] Sheet.
    If [filename] doesn't exist, then this function will create it.

    Parameters:
      filename : File path or existing ExcelWriter
                 (Example: '/path/to/file.xlsx')
      df : dataframe to save to workbook
      sheet_name : Name of sheet which will contain DataFrame.
                   (default: 'Sheet1')
      startrow : upper left cell row to dump data frame.
                 Per default (startrow=None) calculate the last row
                 in the existing DF and write to the next row...
      truncate_sheet : truncate (remove and recreate) [sheet_name]
                       before writing DataFrame to Excel file
      to_excel_kwargs : arguments which will be passed to `DataFrame.to_excel()`
                        [can be dictionary]

    Returns: None
    """
    from openpyxl import load_workbook

    # ignore [engine] parameter if it was passed
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')

    writer = pd.ExcelWriter(filename, engine='openpyxl')

    try:
        # try to open an existing workbook
        writer.book = load_workbook(filename)

        # get the last row in the existing Excel sheet
        # if it was not specified explicitly
        if startrow is None and sheet_name in writer.book.sheetnames:
            startrow = writer.book[sheet_name].max_row

        # truncate sheet
        if truncate_sheet and sheet_name in writer.book.sheetnames:
            # index of [sheet_name] sheet
            idx = writer.book.sheetnames.index(sheet_name)
            # remove [sheet_name]
            writer.book.remove(writer.book.worksheets[idx])
            # create an empty sheet [sheet_name] using old index
            writer.book.create_sheet(sheet_name, idx)

        # copy existing sheets
        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}
    except FileNotFoundError:
        # file does not exist yet, we will create it
        pass

    if startrow is None:
        startrow = 0

    # write out the new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, **to_excel_kwargs)

    # save the workbook
    writer.save()


def series_double_index(dict, df, column_name_1, column_name_2):
    """ This function allows to divide parameters with double index into two dataseries that are added to the
    selected dataframe
    dict: dictionary or parameter. Their keys are in tuple format
    df: dataframe to join the data
    column_name_1: str. Name for the column wih index 1.
    """
    values_1=[]; values_2=[]
    for i, value in enumerate(dict.keys()):
        if i%2 == 0:
            values_1.append(dict[value])
        else:
            values_2.append(dict[value])

    s_1 = pd.Series(values_1)
    s_1.index +=1
    s_2 = pd.Series(values_2)
    s_2.index +=1
    df[column_name_1] = s_1
    df[column_name_2] = s_2

    return df

def write_results_nhl(instance, sheet_name="Scen"):

    #Parameters
    FixedDemandEl_R = instance.FixDemEl.extract_values_sparse()
    MinTemp_R = instance.MinTemp.extract_values_sparse()
    CostEl_R = instance.CEl.extract_values_sparse()
    CostHeat_R = instance.CostHeat.extract_values_sparse()
    Date_R = instance.Date.extract_values_sparse()
    TOut_R = instance.TOut.extract_values_sparse()


    # CapacityHP_R = instance.CapacityHP

    #Variables
        #Main system
    RT401_R = instance.RT401.get_values() # Temperature RT401
    RT506_R = instance.RT506.get_values() # Temperature RT506
    DRQImported_R = instance.DRQImport.get_values() # Temperature transferred to System from DH
        # Building
    Tbinside_R = instance.Tin.get_values() # Temperature inside the bulding
    Thp_R = instance.Thp.get_values()
    RTTbuilding_R = instance.RTTbuilding.get_values() # Temperature transferred to building from System
    socTankQ_R = instance.socTankQ.get_values()
    dischargeTRQ_R = instance.dischargeTRQ.get_values()
    DTQimported_R = instance.DTQimported.get_values()
    Ttank_R = instance.Ttank.get_values()
    TempTank_R = instance.TempTank.get_values()
    DRQ_Need_R = instance.DRQNeeded.get_values()
    dischargeImportTRQ_R = instance.dischargeImportTRQ.get_values()

    TemperatureTank = instance.TempTank[3].value
    Tin_R = instance.Tin[3].value
    RT506_R_1 = instance.RT506[3].value
    RT401_R_1 = instance.RT401[3].value
    maxEl_R = instance.maxEl.value
    maxHeat_R = instance.maxHeat.value


    # Automation of the whole process
    df_auto = []
    periodsi = instance.periods
    for j in periodsi:
        df_auto.append(j)


    df = pd.DataFrame({"Year": Date_R[t],
                       'Minimum Temperature': MinTemp_R[t],
                       'TOut_R': TOut_R[t],
                       'CostEl': CostEl_R[t],
                       'Fixed El Demand': FixedDemandEl_R[t],
                       'Cost Heat': CostHeat_R[t],
                       'MaxHeat':maxHeat_R,
                       'MaxEl': maxEl_R,
                       'RT401': RT401_R[t],
                       "RT506": RT506_R[t],
                       'Thp_R1': Thp_R[t,1],
                       'Thp_R2': Thp_R[t, 2],
                       'socTankQ_R': socTankQ_R[t],
                       'TempTank_R': TempTank_R[t],
                       'dischargeTRQ_R': dischargeTRQ_R[t],
                       'Ttank_R': Ttank_R[t],
                       'DTQimported_R': DTQimported_R[t],
                       'DRQNeeded': DRQ_Need_R[t],
                       'dischargeImportTRQ_R': dischargeImportTRQ_R[t],
                       "Temperature Building":Tbinside_R[t],
                       "RTTbuilding": RTTbuilding_R[t],
                       "DRQImported_R": DRQImported_R[t]} for t in df_auto)

    df10 = pd.DataFrame({"Year": Date_R[t],
                       'Minimum Temperature': MinTemp_R[t],
                       'TOut_R': TOut_R[t],
                       'CostEl': CostEl_R[t],
                       'Fixed El Demand': FixedDemandEl_R[t],
                       'Cost Heat': CostHeat_R[t],
                       'MaxHeat':maxHeat_R,
                       'MaxEl': maxEl_R,
                        'Thp_R1': Thp_R[t, 1],
                        'Thp_R2': Thp_R[t, 2],
                       'RT401': RT401_R[t],
                       "RT506": RT506_R[t],
                       'socTankQ_R': socTankQ_R[t],
                       'TempTank_R': TempTank_R[t],
                       'dischargeTRQ_R': dischargeTRQ_R[t],
                       'Ttank_R': Ttank_R[t],
                       'DTQimported_R': DTQimported_R[t],
                       'DRQNeeded': DRQ_Need_R[t],
                       'dischargeImportTRQ_R': dischargeImportTRQ_R[t],
                       "Temperature Building":Tbinside_R[t],
                       "RTTbuilding": RTTbuilding_R[t],
                       "DRQImported_R": DRQImported_R[t]} for t in range(1,4))



    df2 = pd.DataFrame({"TT": TemperatureTank + random.randint(0,2) - random.randint(0,2)}, index=[0])

    df3 = pd.DataFrame({"TI": Tin_R + round(random.uniform(0.1, 1.0), 10) - round(random.uniform(0.1, 1.0), 10)}, index=[0])

    df4 = pd.DataFrame({"RF506": RT506_R_1}, index=[0])

    df7 = pd.DataFrame({"RF506": RT401_R_1 + random.randint(0, 1) - random.randint(0, 1)}, index=[0])

    df5 = pd.DataFrame({"MaxEl": maxEl_R}, index=[0])

    df6 = pd.DataFrame({"MaxHeat": maxHeat_R}, index=[0])


    append_df_to_excel(r'R_auto_ResTest4_4.xlsx', df, sheet_name=sheet_name)
    append_df_to_excel(r'OperationTest4_4.xlsx', df10, header=None, sheet_name=sheet_name)

    df2.to_excel(r"data/TT.xlsx",index=False, sheet_name="TT")
    df3.to_excel(r"data/TI.xlsx", index=False, sheet_name="TI")
    df4.to_excel(r"data/RF506.xlsx", index=False, sheet_name="RF506")
    df5.to_excel(r"data/MaxEl.xlsx", index=False, sheet_name="MaxEl")
    df6.to_excel(r"data/MaxHeat.xlsx", index=False, sheet_name="MaxHeat")
    df7.to_excel(r"data/RF401.xlsx", index=False, sheet_name="RF401")