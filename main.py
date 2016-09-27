#-*- encoding: utf-8 -*-
'''
Created on 2016年8月25日

@author: BirdLin
'''

from OCVTable import Cell_OCV_Table
import openpyxl
C_mAh = 0
C_Wh = 0
Current_Status =''

#--- Constant Variables -----------------------------------------------
Cell_Series = 10                        # number
Cell_Parallel = 4                       # number
Cell_Capacity = 2600                    # mAh
Design_Capacity = Cell_Parallel * Cell_Capacity
Full_SOC = 100                          # %
Charge_Update_Vol = 39000

Discharge_Current_Threshold = -190       # mA
Charge_Current_Threshold = 100          # mA
Charge_Update_Current_mA = 200          # mA
Charge_Update_Time = 100                 # sec
Static_Update_Time = 80               # sec
Long_Static_Update_Time = 3600        # sec

#--- Tracking & Initial Variables -------------------------------------

Full_Charge_Capacity_mAh = 10400                # mAh

#RM_percentage = (RM_mAh * 100.0 / Full_Charge_Capacity_mAh)
#Used_Capacity_mAh = Full_Charge_Capacity_mAh - RM_mAh ### The capacity has been used.

Discharged_Capacity_calculation = 0           ### Discharged capacity by count minute
Charged_Capacity_calculation = 0              ### Charged capacity by count minute
Extra_Dsg_Cap = 0                               ### Extra discharged capacity

Rest_Time = 0                                    # sec
Rest_Time_for_Update = 0                        # sec
Charge_Time = 0                                 # sec
Charge_Tape_Time = 0                            # sec
Discharge_Time = 0                              # sec

#Static_Update_Enable = False                   # False means can't update or just update; True means available to update.
CurrentInCHG = False
CurrentInDSG = False
CurrentInSTATIC = True
Charge_FCC_Update_Flag = False                 # False is available to update; True means just updated now.
Initialized_Flag = True                        # True means it's in initialized; False means already initialized

### To get static SOC%
def Get_Static_OCV_Percentage():
    i = 0
    global Total_Vol, RM_percentage, Cell_OCV_Table, Cell_Series
    if Total_Vol >= Cell_OCV_Table[0] * Cell_Series :
        RM_percentage = Full_SOC
    elif Total_Vol <= Cell_OCV_Table[99] * Cell_Series :
        RM_percentage = 0
    else :
        while Total_Vol < Cell_OCV_Table[i] * Cell_Series :
            i += 1
            RM_percentage = Full_SOC - i

### To get static capacity
def Normal_Capacity_Calculation():
    global RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, RM_percentage
    Used_Capacity_mAh = (100 - RM_percentage)/100 * Full_Charge_Capacity_mAh
    RM_mAh = Full_Charge_Capacity_mAh - Used_Capacity_mAh
    Discharged_Capacity_calculation = 0
    Charged_Capacity_calculation = 0

def Initialization():
    global Rest_Time_for_Update, Initialized_Flag
    if(Rest_Time_for_Update >= Long_Static_Update_Time or Initialized_Flag == True):
        Get_Static_OCV_Percentage()
        Normal_Capacity_Calculation()
        Initialized_Flag = False
        Rest_Time_for_Update = 0

def Check_Current_Status():
    global Current_mA, CurrentInCHG, CurrentInDSG, CurrentInSTATIC, Current_Status
    if(Current_mA >= Charge_Current_Threshold):
        CurrentInCHG = True
        CurrentInDSG = False
        CurrentInSTATIC = False
        Current_Status = 'Charging'
    elif(Current_mA <= Discharge_Current_Threshold):
        CurrentInCHG = False
        CurrentInDSG = True
        CurrentInSTATIC = False
        Current_Status = 'Discharging'
    else:
        CurrentInCHG = False
        CurrentInDSG = False
        CurrentInSTATIC = True
        Current_Status = 'Rest'

def Calculation_Charging():
    global Current_mA, Charge_Tape_Time, Rest_Time, Rest_Time_for_Update, RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap, Charge_FCC_Update_Flag
    Rest_Time = 0
    Rest_Time_for_Update = 0
    Charged_Capacity_calculation += abs(Current_mA * 10/3600)
    Used_Capacity_mAh += Discharged_Capacity_calculation         ## UC update by discharge capacity
    Used_Capacity_mAh -= Charged_Capacity_calculation            ## UC update by charge capacity
    RM_mAh = Full_Charge_Capacity_mAh - Used_Capacity_mAh
    if(RM_mAh >= Full_Charge_Capacity_mAh):                      ## RM is more than full
        Full_Charge_Capacity_mAh = RM_mAh                        ## FCC update when charging
        RM_percentage = 100
        Used_Capacity_mAh = 0
    elif(RM_mAh < 0):
        RM_mAh = 0
        RM_percentage = 0
    else:                                                        ## RM_mAh >= 0
        RM_percentage = RM_mAh * 100.0 / Full_Charge_Capacity_mAh 
    Charged_Capacity_calculation = 0
    Discharged_Capacity_calculation = 0
    if(Current_mA <= Charge_Update_Current_mA):        ## FCC Charge Update Criteria
        Charge_Tape_Time += 10
        if(Charge_Tape_Time >= Charge_Update_Time and RM_mAh <= Full_Charge_Capacity_mAh and Total_Vol >= Charge_Update_Vol and Charge_FCC_Update_Flag == False):
            Used_Capacity_mAh = 0
            Full_Charge_Capacity_mAh = RM_mAh
            RM_percentage = 100
            Charge_FCC_Update_Flag = True
            Charge_Tape_Time = 0
    Extra_Dsg_Cap = 0                                   ## Only update during discharging and rest

def Calculation_Discharging():
    global Current_mA, Rest_Time, Charge_Tape_Time, Rest_Time_for_Update, RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap, Charge_FCC_Update_Flag
    Rest_Time = 0
    Rest_Time_for_Update = 0
    Charge_Tape_Time = 0
    Discharged_Capacity_calculation += abs(Current_mA * 10/3600)
    Used_Capacity_mAh += Discharged_Capacity_calculation         ## UC update by discharge capacity
    Used_Capacity_mAh -= Charged_Capacity_calculation            ## UC update by charge capacity
    RM_mAh = Full_Charge_Capacity_mAh - Used_Capacity_mAh
    if(RM_mAh >= Full_Charge_Capacity_mAh):
        RM_mAh = Full_Charge_Capacity_mAh                        ## FCC is max RM_mAh when discharging
        RM_percentage = 100
    elif(RM_mAh < 0):
        Used_Capacity_mAh = Full_Charge_Capacity_mAh
        Extra_Dsg_Cap += Discharged_Capacity_calculation         ## count the extra discharge capacity
        RM_mAh = 0
        RM_percentage = 0
    else:                                                        ## RM_mAh >= 0
        RM_percentage = RM_mAh * 100.0 / Full_Charge_Capacity_mAh 
    Charged_Capacity_calculation = 0
    Discharged_Capacity_calculation = 0
    if(RM_percentage <= 70):
        Charge_FCC_Update_Flag = False

def Calculation_Rest():
    global Rest_Time, Rest_Time_for_Update, Charge_Tape_Time, Total_Vol, RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap, Initialized_Flag
    Initialization()
    Charge_Tape_Time = 0                                                ## Need to be adjusted if it's not started in rest
    Rest_Time += 10     
    Rest_Time_for_Update += 10                                             ## count rest time
    ## Static Update
    if(Rest_Time_for_Update >= Static_Update_Time):
        Get_Static_OCV_Percentage()
        if(RM_percentage < 30):
            Used_Capacity_mAh += Extra_Dsg_Cap
            Full_Charge_Capacity_mAh = ((Used_Capacity_mAh*100)/(100-RM_percentage))
            RM_mAh = Full_Charge_Capacity_mAh - Used_Capacity_mAh
            RM_percentage = RM_mAh * 100.0 / (Full_Charge_Capacity_mAh +1)
            Charged_Capacity_calculation = 0
            Discharged_Capacity_calculation = 0
            Extra_Dsg_Cap = 0
            Rest_Time_for_Update = 0
    elif(Rest_Time_for_Update < Static_Update_Time):   
        Used_Capacity_mAh -= Charged_Capacity_calculation            ## UC update by charge capacity
        Used_Capacity_mAh += Discharged_Capacity_calculation         ## UC update by discharge capacity
        RM_mAh = Full_Charge_Capacity_mAh - Used_Capacity_mAh
        if(RM_mAh >= Full_Charge_Capacity_mAh):
            Full_Charge_Capacity_mAh = RM_mAh                        ## FCC update when Charging
            RM_percentage = 100
        elif(RM_mAh < 0):
            Used_Capacity_mAh = Full_Charge_Capacity_mAh
            Extra_Dsg_Cap += Discharged_Capacity_calculation         ## count the extra discharge capacity
            RM_mAh = 0
            RM_percentage = 0
        else:                                                        ## RM_mAh >= 0
            RM_percentage = RM_mAh * 100.0 / Full_Charge_Capacity_mAh 
        Charged_Capacity_calculation = 0
        Discharged_Capacity_calculation = 0


###### Main Gauge Flow ######
#### Reading Data
wb = openpyxl.load_workbook('rest+discharge.xlsx', read_only=False)
ws = wb.get_sheet_by_name ('sheet1')

#### Initialization
## Initialized in rest condition

#### Input Data to Algorithm
row_index = 1
avg_Vol = 0
vol_List = [0] * 10
i = 1

for row in ws.rows:
    if row_index == 1 :
        ws.cell(row=row_index, column=len(row)+1).value='C_mAh'
        ws.cell(row=row_index, column=len(row)+2).value='C_Wh'
        ws.cell(row=row_index, column=len(row)+3).value='mV'
        ws.cell(row=row_index, column=len(row)+4).value='mA'
        ws.cell(row=row_index, column=len(row)+5).value='Current_Status'
        ws.cell(row=row_index, column=len(row)+6).value= 'Rest_Time'
        ws.cell(row=row_index, column=len(row)+7).value= 'Rest_Time_for_Update'
        ws.cell(row=row_index, column=len(row)+8).value= 'Charge_Tape_Time'
        ws.cell(row=row_index, column=len(row)+9).value= 'Charged_Capacity_calculation'
        ws.cell(row=row_index, column=len(row)+10).value= 'Discharged_Capacity_calculation'
        ws.cell(row=row_index, column=len(row)+11).value= 'Extra_Dsg_Cap'
        ws.cell(row=row_index, column=len(row)+12).value= 'Full_Charge_Capacity_mAh'
        ws.cell(row=row_index, column=len(row)+13).value= 'RM_mAh'
        ws.cell(row=row_index, column=len(row)+14).value= 'RM_percentage'
        ws.cell(row=row_index, column=len(row)+15).value= 'Used_Capacity_mAh'
        ws.cell(row=row_index, column=len(row)+16).value= 'Charge_FCC_Update_Flag'
        row_index += 1
        continue
    Total_Vol = row[0].value * 1000
    Current_mA = row[1].value * 1000
    C_mAh += Current_mA * -10 /3600
    C_Wh += Current_mA/-1000 * Total_Vol/1000 * 10 /3600
## Average Voltage
    if i >= 10 :
        i = 10
    j = i-1
    while j > 0 :
        vol_List[j] = vol_List[j-1]
        j -= 1
    vol_List[0] = row[5]
    avg_Vol = sum(vol_List)
    row.append(avg_Vol)
    row.append(avg_Vol/i)
    i+=1
    
#### Current Status Check
    Check_Current_Status()
#### Capacity Calculaton
## Charging                        
    if(CurrentInCHG == True):
        Calculation_Charging()
## Discharging
    elif(CurrentInDSG == True):
        Calculation_Discharging()
## Rest
    elif(CurrentInSTATIC == True):
        Calculation_Rest()
    ws.cell(row=row_index, column=len(row)+1).value = C_mAh
    ws.cell(row=row_index, column=len(row)+2).value = C_Wh    
    ws.cell(row=row_index, column=len(row)+3).value = Total_Vol
    ws.cell(row=row_index, column=len(row)+4).value = Current_mA
    ws.cell(row=row_index, column=len(row)+5).value = Current_Status
    ws.cell(row=row_index, column=len(row)+6).value= Rest_Time
    ws.cell(row=row_index, column=len(row)+7).value= Rest_Time_for_Update
    ws.cell(row=row_index, column=len(row)+8).value= Charge_Tape_Time
    ws.cell(row=row_index, column=len(row)+9).value= Charged_Capacity_calculation
    ws.cell(row=row_index, column=len(row)+10).value= Discharged_Capacity_calculation
    ws.cell(row=row_index, column=len(row)+11).value= Extra_Dsg_Cap
    ws.cell(row=row_index, column=len(row)+12).value= Full_Charge_Capacity_mAh
    ws.cell(row=row_index, column=len(row)+13).value= RM_mAh
    ws.cell(row=row_index, column=len(row)+14).value= RM_percentage
    ws.cell(row=row_index, column=len(row)+15).value= Used_Capacity_mAh
    ws.cell(row=row_index, column=len(row)+16).value= Charge_FCC_Update_Flag


    row_index += 1

#### Save data
print ('Calculation Finished !')
wb.save(filename = 'new_rest+discharge.xlsx')
