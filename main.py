#-*- encoding: utf-8 -*-
'''
Created on 2016年8月25日

@author: BirdLin
'''

from OCVTable import Cell_OCV_Table


#--- Constant Variables -----------------------------------------------
Cell_Series = 10                        # number
Cell_Parallel = 4                       # number
Cell_Capacity = 2600                    # mAh
Design_Capacity = Cell_Parallel * Cell_Capacity
Full_SOC = 100                          # %
Charge_Update_Vol = 39000

Discharge_Current_Threshold = -200       # mA
Charge_Current_Threshold = 100          # mA
Charge_Update_Current_mA = 200          # mA
Charge_Update_Time = 60                 # sec
Static_Update_Time = 5400               # sec


#--- Tracking & Initial Variables -------------------------------------

Full_Charge_Capacity_mAh = 10400                # mAh
RM_mAh = 1000                                   # mAh
RM_percentage = (RM_mAh * 100.0 / Full_Charge_Capacity_mAh)
Used_Capacity_mAh = Full_Charge_Capacity_mAh - RM_mAh ### The capacity has been used.

Total_Vol = 34000
Static_Vol = 34000                              # mV
Discharged_Capacity_calculation = 100           ### Discharged capacity by count minute
Charged_Capacity_calculation = 400              ### Charged capacity by count minute
Extra_Dsg_Cap = 0                               ### Extra discharged capacity

Static_OCV_Percentage = 0                       # %
Rest_Time = 5400                                # sec
Charge_Time = 0                                 # sec
Charge_Tape_Time = 0                            # sec
Discharge_Time = 0                              # sec

Current_mA = 50                                # mA
Static_Update_Enable = False
CurrentInCHG = False
CurrentInDSG = False
CurrentInSTATIC = True
Charge_FCC_Update_Flag = False

### To get static SOC%
def Get_Static_OCV_Percentage(Static_Vol, Cell_OCV_Table, Cell_Series):
    i = 0
    global Static_OCV_Percentage
    if Static_Vol >= Cell_OCV_Table[0] * Cell_Series :
        Static_OCV_Percentage = Full_SOC
    elif Static_Vol <= Cell_OCV_Table[99] * Cell_Series :
        Static_OCV_Percentage = 0
    else :
        while Static_Vol < Cell_OCV_Table[i] * Cell_Series :
            i += 1
            Static_OCV_Percentage = Full_SOC - i


### To get static capacity
def Initial_Static_Capacity_Calculation(Static_OCV_Percentage):
    global RM_mAh, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, RM_percentage
    RM_mAh = (Static_OCV_Percentage / 100.0 * Full_Charge_Capacity_mAh)
    RM_percentage = RM_mAh * 100.0 / Full_Charge_Capacity_mAh
    Used_Capacity_mAh = Full_Charge_Capacity_mAh - RM_mAh
    Discharged_Capacity_calculation = 100
    Charged_Capacity_calculation = 200


print ('Current I = %d mA, V = %d mV, RM_Percent = %5.1f' % (Current_mA, Static_Vol, Static_OCV_Percentage))

### Initialization
Get_Static_OCV_Percentage(Static_Vol, Cell_OCV_Table, Cell_Series)
Initial_Static_Capacity_Calculation(Static_OCV_Percentage)
print ('Initialization : RM_mAh:%5d  RM_Percent:%5.1f  UC:%5d  DCC:%5d  CCC:%5d  FCC:%5d  Extra_Cap:%5d' % (RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap))

#print 'RM_mAh:%s UC:%s DCC:%s CCC:%d RM%:%s FCC:%s' % (RM_mAh, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, RM_percentage, Full_Charge_Capacity_mAh)

def Check_Current_Status(Current_mA):
    global CurrentInCHG, CurrentInDSG, CurrentInSTATIC
    if(Current_mA >= Charge_Current_Threshold):
        CurrentInCHG = True
        CurrentInDSG = False
        CurrentInSTATIC = False
        print ('Charging')
    elif(Current_mA <= Discharge_Current_Threshold):
        CurrentInCHG = False
        CurrentInDSG = True
        CurrentInSTATIC = False
        print ('Discharging')
    else:
        CurrentInCHG = False
        CurrentInDSG = False
        CurrentInSTATIC = True
        print ('Rest')

def Calculation_Charging():
    global Rest_Time, RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap, Charge_FCC_Update_Flag
    Rest_Time = 0
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
        Charge_Tape_Time += 1
        if(Charge_Tape_Time >= Charge_Update_Time and RM_mAh <= Full_Charge_Capacity_mAh and Total_Vol >= Charge_Update_Vol and Charge_FCC_Update_Flag == False):
            Used_Capacity_mAh = 0
            Full_Charge_Capacity_mAh = RM_mAh
            RM_percentage = 100
            Charge_FCC_Update_Flag = True
    Extra_Dsg_Cap = 0                                   ## Only update during discharging and rest
    print ('Charging : RM_mAh:%5d  RM_Percent:%5.1f  UC:%5d  DCC:%5d  CCC:%5d  FCC:%5d  Extra_Cap:%5d' % (RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap))

def Calculation_Discharging():
    global Rest_Time, RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap, Charge_FCC_Update_Flag
    Rest_Time = 0
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
    if(RM_percentage <= 70 and Charge_FCC_Update_Flag == True):
        Charge_FCC_Update_Flag = False
    print ('Discharging : RM_mAh:%5d  RM_Percent:%5.1f  UC:%5d  DCC:%5d  CCC:%5d  FCC:%5d  Extra_Cap:%5d' % (RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap))

def Calculation_Rest():
    global Rest_Time, RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap
    Rest_Time += 1                                                  ## count rest time
    ## Static Update
    if(Rest_Time >= Static_Update_Time):
        Get_Static_OCV_Percentage(Static_Vol, Cell_OCV_Table, Cell_Series)
        if(Static_OCV_Percentage < 30):
            Extra_Dsg_Cap = 100                                     ## for test
            Used_Capacity_mAh += Extra_Dsg_Cap
            Full_Charge_Capacity_mAh = ((Used_Capacity_mAh*100)/(100-Static_OCV_Percentage))
            RM_mAh = Full_Charge_Capacity_mAh - Used_Capacity_mAh
            RM_percentage = RM_mAh * 100.0 / Full_Charge_Capacity_mAh
            Charged_Capacity_calculation = 0
            Discharged_Capacity_calculation = 0
            Extra_Dsg_Cap = 0
    elif(Rest_Time < Static_Update_Time):   
        Used_Capacity_mAh -= Charged_Capacity_calculation            ## UC update by charge capacity
        Used_Capacity_mAh += Discharged_Capacity_calculation         ## UC update by discharge capacity
        #Used_Capacity_mAh = 10500
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
    print ('Rest : RM_mAh:%5d  RM_Percent:%5.1f  UC:%5d  DCC:%5d  CCC:%5d  FCC:%5d  Extra_Cap:%5d' % (RM_mAh, RM_percentage, Used_Capacity_mAh, Discharged_Capacity_calculation, Charged_Capacity_calculation, Full_Charge_Capacity_mAh, Extra_Dsg_Cap))



###### Main Gauge Flow ######
## Current Status Check
Check_Current_Status(Current_mA)


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


