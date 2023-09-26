# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 23:37:58 2022

@author: schiaffipn
Optimization Model for Energy Hub. 
Six Decision Variables: Number of Turbines, Solar PV Area, Storage Size (Batteries and Hydrogen storage), Installed capacity of Electrolyser & Compressor 
Assumptions and adopted values:
    - Locations
    - Distances to shore
    - Electricity and H2 price
"""
import pandas as pd
import numpy as np
import pyomo.environ as pyo
import matplotlib.pyplot as plt
import os
import numpy_financial as npf

"--------------------Reading Input Data-------------------"
"---------------------------------------------------"
#Read Input File
cwd = os.getcwd()                                                              #get Current Working Directory        
input_file = os.path.join(cwd, 'input', "input_file.xlsx")                     #create Input File directory

def readExcel (input_file):                                                    #Reading Excel Sheets, use index_col argunment to have the first column as the value of row, then I use loc
 
    df_general = pd.read_excel(input_file,sheet_name="General", header = 0, usecols=[0,1,2],index_col=(0))
    df_economic = pd.read_excel(input_file,sheet_name="Economic", header = 0, usecols=[0,1,2],index_col=(0))
    df_solar = pd.read_excel(input_file, sheet_name="Solar", header = 0, usecols=[0,1,2], index_col=(0))
    df_wind = pd.read_excel(input_file,sheet_name="Wind", header = 0, usecols=[0,1], index_col=(0))
    df_storage = pd.read_excel(input_file, sheet_name="Storage", header = 0, usecols=[0,1,2], index_col=(0))
    df_Hydrogen = pd.read_excel(input_file, sheet_name="Hydrogen", header = 0, usecols=[0,1,2], index_col=(0))
    df_EWind_h = pd.read_excel(input_file,sheet_name="Wind_Power_Data", header = 0, usecols=[0,1,2])
    df_EPV_h = pd.read_excel(input_file,sheet_name="Solar_Power_Data", header = 0, usecols=[0,1,2])
    df_electricity_prices = pd.read_excel(input_file,sheet_name="Day-ahead Prices_2015-2022", header = 0)
    
    return df_general, df_economic, df_solar, df_wind, df_storage, df_Hydrogen, df_EWind_h, df_EPV_h, df_electricity_prices

df_general, df_economic, df_solar, df_wind, df_storage, df_Hydrogen, df_EWind_h, df_EPV_h, df_electricity_prices = readExcel(input_file)

print('\n Energy HUB Model North Sea')
print('\n Shore Distance [km]', df_general.loc['shore_distance','Input'])
print('\n Area Hub [km2]', df_general.loc['area_hub','Input'])


"np_calculator(C,r) Function that allows me to get the NPV of a costs or production flow given the rate of return"
"C = the costs or energy production flow, it is a list, or df, indexed over a lifetime (years)"
"r = rate of return "
def np_calculator(C,r): # net present calculator
    net_present = 0
    for k in range(C.size): #range (0 to lifetime)
        net_present = net_present + C[k]/(1+r)**k
    return net_present

"This function creates a flow(array) of costs or energy production using a np array, starting in year 0"
def costs_production_flow(opex_production_asset,lifetime):
    cost_production_flow = np.zeros(lifetime+1) #"OPEX or Production start from year 1 to lifetime --> therefore +1, in year 0 is reserved for CAPEX and Production=0
    cost_production_flow[1:lifetime+1] = opex_production_asset #opex or production starts from year 1.. 
    return cost_production_flow

def power_compressor(df_Hydrogen):
    # Calculate estimated compressor power
    Tmean = 333.15  # Inlet temperature of the compressor
    gamma = 1.4  # Specific heat ratio
    #P_out = Electrolyser['pipeline_inlet']  # Compressor output pressure
    Pout = df_Hydrogen.loc['export_pressure','Input']*100   #bar to kPa
    Pin =  df_Hydrogen.loc['output_pressure','Input']*100   #bar to kPa
    GH = 0.0696
    P_compressor = ((286.76/(GH*0.85*0.98))*Tmean*(gamma/(gamma-1))*(((Pout/Pin)**((gamma-1)/gamma))-1))/3600000     #J to kWh
    return P_compressor

compressor_power_94bar = power_compressor(df_Hydrogen)    # for current user input Pin, Pout (94 bar for 85km with on shore oressure of 68bar)

def power_compressor():
    # Calculate estimated compressor power
    Tmean = 333.15  # Inlet temperature of the compressor
    gamma = 1.4  # Specific heat ratio
    #P_out = Electrolyser['pipeline_inlet']  # Compressor output pressure
    Pout = 200*100   #bar to kPa
    Pin =  30*100   #bar to kPa
    GH = 0.0696
    P_compressor = ((286.76/(GH*0.85*0.98))*Tmean*(gamma/(gamma-1))*(((Pout/Pin)**((gamma-1)/gamma))-1))/3600000     #J to kWh
    return P_compressor

compressor_power_200bar = power_compressor()    # for current user input Pin, Pout

#Calculations for the NPC to be used in the Model
#1st Create a cost flow of the OPEX
#2nd If lifetime of asseet is lower than lifetime of the project, then add the corresponing extra costs in the correct year
#3nd Calculate NPC

PV_OPEX_flow = costs_production_flow(df_economic.loc['OPEX_solar_total','Input'],int(df_general.loc['system_lifetime','Input'])) 
np_PV_OPEX = np_calculator(PV_OPEX_flow,df_general.loc['discount_rate','Input'])                                  

Wind_OPEX_flow = costs_production_flow(df_economic.loc['OPEX_wind','Input'],int(df_general.loc['system_lifetime','Input'])) 
np_WIND_OPEX= np_calculator(Wind_OPEX_flow,df_general.loc['discount_rate','Input']) 

Storage_OPEX_flow = costs_production_flow(df_economic.loc['OPEX_battery','Input'],int(df_general.loc['system_lifetime','Input']))
np_Storage_OPEX = np_calculator(Storage_OPEX_flow,df_general.loc['discount_rate','Input'])

Electrolyser_OPEX_flow = costs_production_flow(df_economic.loc['OPEX_Electrolysis','Input'],int(df_general.loc['system_lifetime','Input']))
#Adding Investment costs of the stack over the lifetime
i=1
while df_Hydrogen.loc['lifetime_stack','Input']*i <= df_general.loc['system_lifetime','Input']:
    Electrolyser_OPEX_flow[df_Hydrogen.loc['lifetime_stack','Input']*i]=Electrolyser_OPEX_flow[df_Hydrogen.loc['lifetime_stack','Input']*i] + df_Hydrogen.loc['CAPEX_stack','Input']
    i = i + 1
np_Electrolyser_OPEX = np_calculator(Electrolyser_OPEX_flow,df_general.loc['discount_rate','Input'])

 
Compressor_OPEX_flow = costs_production_flow(df_economic.loc['OPEX_compressor','Input'],int(df_general.loc['system_lifetime','Input']))
np_Compressor_OPEX = np_calculator(Compressor_OPEX_flow, df_general.loc['discount_rate','Input'])

H2_Storage_flow =  costs_production_flow(df_economic.loc['OPEX_H2_storage','Input'],int(df_general.loc['system_lifetime','Input']))
np_H2_Storage_OPEX = np_calculator(H2_Storage_flow, df_general.loc['discount_rate','Input'])

H2_pipe = costs_production_flow(df_Hydrogen.loc['OPEX_pipe','Input'],int(df_general.loc['system_lifetime','Input']))
np_H2_pipe = np_calculator(H2_pipe, df_general.loc['discount_rate','Input'])

"---------------------------------PYOMO MODEL---------------------------------"
def CreateModel (df_general, df_economic, df_solar, df_wind, df_storage, df_Hydrogen, df_EWind_h, df_EPV_h,df_electricity_prices):
    model = pyo.ConcreteModel(name='HPP - Model Optimisation')
    
    #Defining Sets
    # model.T = pyo.Set(ordered = True, initialize=range(8760)) ATTENTION: RangeSet starts in 1!, Range starts in 0! 
    model.T = pyo.Set(initialize = pyo.RangeSet(len(df_EPV_h)), ordered=True) 
    
    "----------------------Parameters----------------------"
    ##Solar
    model.PV_CAPEX = pyo.Param(initialize=df_economic.loc['CAPEX_FPV','Input'], mutable=(True))
    model.PV_OPEX = pyo.Param(initialize=np_PV_OPEX)
    
    ##Wind Costs
    model.W_CAPEX = pyo.Param(initialize=df_economic.loc['CAPEX_wind','Input'], mutable=(True))
    model.W_OPEX = pyo.Param(initialize=np_WIND_OPEX)
    
    #Storage 
    model.CAPEX_Storage = pyo.Param(initialize=(df_economic.loc['CAPEX_battery','Input']), mutable=(True))
    model.OPEX_Storage = pyo.Param(initialize=np_Storage_OPEX)
    model.charge_rate = pyo.Param(initialize=df_storage.loc['charge_rate','Input'], mutable=(True))
    model.discharge_rate = pyo.Param(initialize=df_storage.loc['discharge_rate','Input'], mutable=(True))   
    
    ##Electrolyser
    model.CAPEX_electrolyser = pyo.Param(initialize=(df_economic.loc['CAPEX_Electrolysis','Input']), mutable=(True))
    model.OPEX_electrolyser = pyo.Param(initialize=np_Electrolyser_OPEX)
    model.electrolyser_efficiency = pyo.Param(initialize=(df_Hydrogen.loc['Electrolysis_Efficiency','Input']), mutable=(True)) #kWh/kg
    
    #Compressor
    model.CAPEX_compressor = pyo.Param(initialize = df_economic.loc['CAPEX_compressor','Input'])
    model.OPEX_compressor = pyo.Param(initialize = np_Compressor_OPEX)
    model.compressor_power_94bar = pyo.Param(initialize = compressor_power_94bar)
    model.compressor_power_200bar = pyo.Param(initialize = compressor_power_200bar)
    #H2 Storage
    model.CAPEX_h2_storage = pyo.Param(initialize=df_economic.loc['CAPEX_H2_storage','Input'])
    model.OPEX_h2_storage = pyo.Param(initialize = np_H2_Storage_OPEX)
    
    #H2 Pipe
    model.H2_Pipe =  pyo.Param(initialize = np_H2_pipe) #the opex of teh pipe NPC

    #Hydrogen Price EUR/kg
    model.H2_price = pyo.Param(initialize = 10, mutable=(True))
    
    
    "--------------------Decision Variables--------------------"
    #Decision Variables
    model.x1 = pyo.Var(within=pyo.NonNegativeReals)    # Area PV [m2]
    model.x2 = pyo.Var(within=pyo.NonNegativeIntegers) # Number of Wind Turbines
    model.x3 = pyo.Var(within=pyo.NonNegativeReals)    # Storage Size [kWh]
    model.x4 = pyo.Var(within=pyo.NonNegativeReals)    # Electrolyser Size [kWh]                       
    model.x5 = pyo.Var(within=pyo.NonNegativeReals)    # Compressor Size  power for exporting from 30 bar to export pressure in pippe
    model.x5b = pyo.Var(within=pyo.NonNegativeReals)    # Compressor Size required for 200 bar storage
    model.x6 = pyo.Var(within=pyo.NonNegativeReals, bounds=(0,3000000))    # h2 buffer [kg]
    
    "------------------------Variables - the Flows------------------------"
    
    model.EW_h = pyo.Var(model.T,within=pyo.NonNegativeReals)    # Total Wind Energy                 
    model.EPV_h = pyo.Var(model.T,within=pyo.NonNegativeReals)   # Total PV Energy      
    
    model.EUsed_h = pyo.Var(model.T,within=pyo.NonNegativeReals)    #Total Electricity used
    model.ECurtailed_h = pyo.Var(model.T,within=pyo.NonNegativeReals)   #Total electricity Curtailed
    
    model.EUsed_Grid_h = pyo.Var(model.T,within=pyo.NonNegativeReals)        #Electricity Used Wind and PV to Grid  
    model.EUsed_Battery_h = pyo.Var(model.T,within=pyo.NonNegativeReals)     #Electricity from Wind and PV to Storage (Charge_Flow)      
    model.EUsed_H2_h = pyo.Var(model.T,within=pyo.NonNegativeReals)             #Electriity used from W and PV to H2 production 
   
    model.EH2_Electrolyser_h = pyo.Var(model.T,within=pyo.NonNegativeReals)     #Electricity from Wind and PV to Electrolyser 
    model.EH2_Compressor_h = pyo.Var(model.T,within=pyo.NonNegativeReals)       #Electricity from Wind and PV to Compressor 
    
    model.EBattery_Grid_h = pyo.Var(model.T, within=pyo.NonNegativeReals)           #Electricity from baterry to grid (discharge flow)
    model.EBattery_Electrolyser_h = pyo.Var(model.T, within=pyo.NonNegativeReals)   #Electricity from baterry to electrolyser (discharge flow)
    model.PIPEEBattery_Compressor_h = pyo.Var(model.T, within=pyo.NonNegativeReals)
    model.STOEBattery_Compressor_h = pyo.Var(model.T, within=pyo.NonNegativeReals)
    
    model.SoC_h = pyo.Var(model.T, within=pyo.NonNegativeReals)                 #Battery SoC
    
    model.EElectrolyser_h = pyo.Var(model.T,within=pyo.NonNegativeReals)        #Total electricity used from PV, Wind and Storage! 
    model.PIPEECompressor_h = pyo.Var(model.T,within=pyo.NonNegativeReals)          #Total electricity used from PV, Wind and Storage!
    model.STOECompressor_h = pyo.Var(model.T,within=pyo.NonNegativeReals)  
    model.ECompressor_PIPE_h = pyo.Var(model.T,within=pyo.NonNegativeReals) 
    model.ECompressor_STO_h = pyo.Var(model.T,within=pyo.NonNegativeReals) 
    
    model.E_H2_Total_h = pyo.Var(model.T,within=pyo.NonNegativeReals)           
    model.H2_flow_h = pyo.Var(model.T,within=pyo.NonNegativeReals)              #Total H2 production per hour from the Electrolyser
    model.H2_EL_STO_h = pyo.Var(model.T,within=pyo.NonNegativeReals)            #H2 from Electrolyser to storage  
    model.H2_EL_PIPE_h = pyo.Var(model.T,within=pyo.NonNegativeReals)            #H2 from electrolyser to CMP--> to pipe
    model.H2_STO_PIPE_h = pyo.Var(model.T,within=pyo.NonNegativeReals)           #H2 from Buffer/Storage to pipe
    model.H2_Export_h = pyo.Var(model.T,within=pyo.NonNegativeReals)

    model.h2Storage = pyo.Var(model.T,within=pyo.NonNegativeReals)              # H2 Storage
    model.Electricity_export = pyo.Var(model.T,within=pyo.NonNegativeReals)  

    model.variation = pyo.Var(model.T,within=pyo.Reals) 
    
# OBJECTIVE FUNCTION
    def OF (model):                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            #+sum((model.E_H2_Total_h[t]*df_electricity_prices.loc[t-1,'Average_2015_2022_[EUR/kWh]'])/((1+df_general.loc['discount_rate','Input'])**(i+1)) for t in model.T for i in range(int(df_general.loc['system_lifetime','Input'])))                                                                                                                                                                                                                                                                                                                                                                                                                                                         
        return sum((model.H2_Export_h[t]*model.H2_price)/((1+df_general.loc['discount_rate','Input'])**(i+1)) for t in model.T for i in range(int(df_general.loc['system_lifetime','Input']))) + sum((model.Electricity_export[t]*df_electricity_prices.loc[t-1,'Average_2015_2022_[EUR/kWh]'])/((1+df_general.loc['discount_rate','Input'])**(i+1)) for t in model.T for i in range(int(df_general.loc['system_lifetime','Input']))) - (((model.x1*df_solar.loc['FPV_kWpm2','Input'])*(model.PV_CAPEX+model.PV_OPEX))  + (model.x2*(model.W_CAPEX+model.W_OPEX)) + (model.x3*(model.CAPEX_Storage + model.OPEX_Storage)) + (model.x4*(model.CAPEX_electrolyser + model.OPEX_electrolyser)) + ((model.x5+model.x5b)*(model.CAPEX_compressor + model.OPEX_compressor)) + (model.x6*(model.CAPEX_h2_storage + model.OPEX_h2_storage)) +  model.H2_Pipe + df_Hydrogen.loc['CAPEX_pipe','Input']) #
    model.ObjFunction = pyo.Objective(rule=OF, sense = pyo.maximize)

# Electricity Equations      
    #"EPV_h: Total energy generated by PV"    
    def electricity1 (model,t):
        return model.EPV_h[t] == model.x1*df_EPV_h.loc[t-1,'Power_Output[kWh/m2]']
    #"EW_h: Total energy generated by Wind"
    def electricity2 (model,t):
        return model.EW_h[t] == model.x2*df_EWind_h.loc[t-1,'Wind_Power_1']
    #"Total Used energy is equal to the total generated by PV and Wind minus the energy cutailed"
    def electricity3 (model,t):
        return model.EUsed_h[t] == (model.EPV_h[t] + model.EW_h[t]) - model.ECurtailed_h[t] 
    #"Total Used energy is equal to energy used to the Grid, plus Battery plus Hydrogen production"
    def electricity4 (model,t):
        return model.EUsed_h[t] == model.EUsed_Grid_h[t] + model.EUsed_Battery_h[t] + model.EUsed_H2_h[t]         
      
#SoC Equations 
    #"SoC1: At every hour h, the storage is equal to the storage level at the previous hour plus te charging flows multipliied by the charge rate miinus the discharging flows"
    def SoC1 (model,t):
        # if t == 0:
        if t == 1:
            return  model.SoC_h[t] == 0 # if initial storage is not enough for the sum of the demand in the first hours without irradiance, then the model does not work! in this case the sum is 136.9 kWh #+ (model.charge_flow[t]*df_technical_data.loc['charge_rate','Input']) - (model.discharge_flow[t]/df_technical_data.loc['discharge_rate','Input'])
        else:
            return  model.SoC_h[t] ==  (model.SoC_h[t-1]*df_storage.loc['Self_discharge','Input']) + (model.EUsed_Battery_h[t]*model.charge_rate) - ((model.EBattery_Grid_h[t]+model.EBattery_Electrolyser_h[t]+model.PIPEEBattery_Compressor_h[t]+model.STOEBattery_Compressor_h[t])/model.discharge_rate)

#"SoC2: At every hour the storage level has to be lower or equal than the maxium (X3)"
# To avoid degradation, teh DoD is 80%, then the battery opperates between 90% and 10% of total capacity
    def SoC2 (model, t):
        return model.SoC_h[t] <= model.x3*0.90    
    def SoC3 (model, t):
        if t == 1:
            return model.SoC_h[t] == 0
        else:
            return model.SoC_h[t] >= model.x3*0.10 

    
    #"SoC3: The discharge flow cannot be higher than the storage level at h-1 (previous hour)"    
    def SoC4 (model, t):
        if t == 1:
            return model.SoC_h[t] == 0
        else:
            return (model.EBattery_Grid_h[t] + model.EBattery_Electrolyser_h[t] + model.PIPEEBattery_Compressor_h[t]+model.STOEBattery_Compressor_h[t])/model.discharge_rate <= model.SoC_h[t-1]
    #"SoC4: The charge flow cannot be higher than the the maximum capaacity minus SoC prevoius hour"    
    def SoC5 (model, t):
        if t == 1:
            return model.SoC_h[t] == 0
        else:
            return model.EUsed_Battery_h[t]*model.charge_rate <= model.x3*0.90 - model.SoC_h[t-1]    
    def SoC6 (model, t):
        return (model.EUsed_Battery_h[t]*model.charge_rate) <= model.x3*df_storage.loc['charge_discharge_power','Input']
    def SoC7 (model, t):
        return ((model.EBattery_Grid_h[t] + model.EBattery_Electrolyser_h[t] + model.PIPEEBattery_Compressor_h[t] + model.STOEBattery_Compressor_h[t])/model.discharge_rate) <= model.x3*df_storage.loc['charge_discharge_power','Input']
 
            
#Balance Electricity 
    def balance0 (model,t):    
        return model.Electricity_export[t] == model.EUsed_Grid_h[t] + model.EBattery_Grid_h[t]    

# Hydrogen Equations
# Total Electricity to Hydrogen production goes to Electrolyser or Compressor
    def hydrogen0 (model, t):
        return  model.EUsed_H2_h[t] == model.EH2_Electrolyser_h[t] + model.EH2_Compressor_h[t]

# Electrolyser Equations
# Total Electricity Consumption Electrolyser
    def electrolyser0 (model, t):
        return  model.EElectrolyser_h[t] == model.EH2_Electrolyser_h[t] + model.EBattery_Electrolyser_h[t]


########### OPTIONAL ##########
# Management electricity to electrolyser equations, limit the diffefrent in load between hours. 
    def electrolyser1 (model, t):
        if t == 1:
            return model.variation[t]==0
        else:
            return model.variation[t] == model.EElectrolyser_h[t] - model.EElectrolyser_h[t-1]        
    def electrolyser2 (model, t):
        if t == 1:
            return model.variation[t]==0
        else:
            return model.variation[t] <= 10 
    def electrolyser3 (model, t):
        if t == 1:
            return model.variation[t]==0
        else:
            return model.variation[t] >= -10         
########### OPTIONAL ##########  

# Total Hydrogen Production
    def electrolyser4 (model, t):
        return model.H2_flow_h[t] == model.EElectrolyser_h[t]/model.electrolyser_efficiency #kg H2
# Total H2 produced goes to the storage meduim or to the compressor to be exported
    def electrolyser5 (model,t):
        return model.H2_flow_h[t] == model.H2_EL_STO_h[t] + model.H2_EL_PIPE_h[t] 
# Electrolyser Capacity
    def electrolyser6 (model, t):
        return model.EElectrolyser_h[t] <= model.x4

#Hydrogen Storage
    def h2Storage1 (model, t):
        if t == 1:
            return model.h2Storage[t] == 0
        else:
            return model.h2Storage[t] == model.h2Storage[t-1] + model.H2_EL_STO_h[t] - model.H2_STO_PIPE_h[t]
    def h2Storage2 (model, t):
        return model.h2Storage[t] <= model.x6
    def h2Storage3 (model, t):
        if t == 1:
            return model.h2Storage[t] == 0
        else:
            return model.H2_STO_PIPE_h[t] <= model.h2Storage[t-1]
    def h2Storage4 (model, t):
        if t == 1:
            return model.h2Storage[t] == 0
        else:
            return model.H2_EL_STO_h[t] <= model.x6 - model.h2Storage[t-1]        

# Balance Hydrogen    
    def balanceH21 (model,t):
        return model.H2_Export_h[t] == (model.H2_EL_PIPE_h[t] + model.H2_STO_PIPE_h[t]) 
    def balanceH22 (model,t):
        return model.H2_Export_h[t] <= 0.85*df_Hydrogen.loc['pipe_capacity','Input'] #max utilisation is 85%
   
    
   
# compressor Pipe
# Compressor Equations
    def compressor0 (model, t):
        return model.EH2_Compressor_h[t] ==  model.PIPEECompressor_h[t] + model.STOECompressor_h[t] #electricity from PV and wind goind directly to compressor

    def compressor4 (model, t):
       return  model.ECompressor_PIPE_h[t] == model.PIPEECompressor_h[t] + model.PIPEEBattery_Compressor_h[t] #total electricity consumed to compress H2 to pipe, it comes from PV+W directly + battery
 
    def compressor5 (model, t):
        return  model.ECompressor_STO_h[t] ==  model.STOECompressor_h[t] + model.STOEBattery_Compressor_h[t] # same but for storage H2
    
    # def compressor1 (model, t):
    #     return  model.PIPEECompressor_h[t] == (model.H2_EL_PIPE_h[t]*model.compressor_power_94bar)
    def compressor1 (model, t):
        return  model.ECompressor_PIPE_h[t] == (model.H2_EL_PIPE_h[t]*model.compressor_power_94bar)
  
    
#  Electricity Consumption Compressor for hydrogen from electrolyser to STORAGE directly"
    # def compressor2 (model, t):
    #     return  model.STOECompressor_h[t] ==  (model.H2_EL_STO_h[t]*model.compressor_power_200bar) 
    def compressor2 (model, t):
        return  model.ECompressor_STO_h[t] ==  (model.H2_EL_STO_h[t]*model.compressor_power_200bar)     
# Compressor Capacity"
    def compressor3 (model, t):
        return model.ECompressor_PIPE_h[t] <= model.x5

# Compressor Capacity"
    def compressor6 (model, t):
        return model.ECompressor_STO_h[t] <= model.x5b

#Total electricity consumption Hydrogen production
    def H2_electricity_total (model,t):
        return model.E_H2_Total_h[t] == model.ECompressor_PIPE_h[t] +model.ECompressor_STO_h[t]+ model.EElectrolyser_h[t] 

#Area constraint
    def area_check (model,t):
        return (((model.x1*df_solar.loc['FPV_kWpm2','Input']*df_solar.loc['area_FPV','Input'])/1000000) + (model.x2*df_wind.loc['required_area_turbine','Input'])) <= 0.9*df_general.loc['area_hub','Input']



    model.celectricity1 = pyo.Constraint(model.T, rule=electricity1)
    model.celectricity2 = pyo.Constraint(model.T, rule=electricity2)
    model.celectricity3 = pyo.Constraint(model.T, rule=electricity3)
    model.celectricity4 = pyo.Constraint(model.T, rule=electricity4)
    model.cSoC1 = pyo.Constraint(model.T, rule=SoC1)
    model.cSoC2 = pyo.Constraint(model.T, rule=SoC2)
    model.cSoC3 = pyo.Constraint(model.T, rule=SoC3)
    model.cSoC4 = pyo.Constraint(model.T, rule=SoC4)
    model.cSoC5 = pyo.Constraint(model.T, rule=SoC5)
    model.cSoC6 = pyo.Constraint(model.T, rule=SoC6)        
    model.cSoC7 = pyo.Constraint(model.T, rule=SoC7) 
    model.cbalance0 = pyo.Constraint(model.T, rule=balance0)
    model.chydrogen0 = pyo.Constraint(model.T, rule=hydrogen0)
    model.celectrolyser0 = pyo.Constraint(model.T, rule=electrolyser0)
    ########### OPTIONAL ##########
    model.celectrolyser1 = pyo.Constraint(model.T, rule=electrolyser1) 
    model.celectrolyser2 = pyo.Constraint(model.T, rule=electrolyser2)    
    model.celectrolyser3 = pyo.Constraint(model.T, rule=electrolyser3)   
    ########### OPTIONAL ##########
    model.celectrolyser4 = pyo.Constraint(model.T, rule=electrolyser4)
    model.celectrolyser5 = pyo.Constraint(model.T, rule=electrolyser5) 
    model.celectrolyser6 = pyo.Constraint(model.T, rule=electrolyser6)
    model.ch2Storage1 = pyo.Constraint(model.T, rule=h2Storage1) 
    model.ch2Storage2 = pyo.Constraint(model.T, rule=h2Storage2)
    model.ch2Storage3 = pyo.Constraint(model.T, rule=h2Storage3)
    model.ch2Storage4 = pyo.Constraint(model.T, rule=h2Storage4)
    model.ccompressor0 = pyo.Constraint(model.T, rule=compressor0)
    model.ccompressor1 = pyo.Constraint(model.T, rule=compressor1)
    model.ccompressor2 = pyo.Constraint(model.T, rule=compressor2)
    model.ccompressor3 = pyo.Constraint(model.T, rule=compressor3)
    model.ccompressor4 = pyo.Constraint(model.T, rule=compressor4)    
    model.ccompressor5 = pyo.Constraint(model.T, rule=compressor5)
    model.ccompressor6 = pyo.Constraint(model.T, rule=compressor6)
    model.cbalanceH21 = pyo.Constraint(model.T, rule=balanceH21)
    model.cbalanceH22 = pyo.Constraint(model.T, rule=balanceH22)
    model.cH2_electricity_total = pyo.Constraint(model.T, rule=H2_electricity_total)
    model.carea_check = pyo.Constraint(model.T, rule=area_check)

    
    return model

model = CreateModel(df_general, df_economic, df_solar, df_wind, df_storage, df_Hydrogen, df_EWind_h, df_EPV_h,df_electricity_prices)      
  
opt = pyo.SolverFactory('gurobi')
results = opt.solve(model, tee=True)
# model.pprint()
results.write()
# model.display()


installed_power_PV = pyo.value(model.x1)*df_solar.loc['FPV_kWpm2','Input']

# print('\n FLH = ', pyo.value(model.FLH))

print('\n ---------------------------------------------------')
print('\n Decision Variables: ')
print('\n Total area FPV = %4.2f [km2]' %((installed_power_PV*df_solar.loc['area_FPV','Input'])/1000000))
print('\n Total number of turbines = ', pyo.value(model.x2))
print('\n Total area wind farm = %4.2f [km2]' %(pyo.value(model.x2)*df_wind.loc['required_area_turbine','Input']))
print('\n Total Baterry capacity [kWh] = ', pyo.value(model.x3))
print('\n Electrolyser Capacity [kW] = ', pyo.value(model.x4))
print('\n Compressor Capacity [kW]= ', pyo.value(model.x5)+pyo.value(model.x5b))
print('\n H2 Storage Capacity - Salt Caverns [kg]= ', pyo.value(model.x6))    


print('\n ---------------------------------------------------')
print('\n Results: ')

total_area = ((installed_power_PV*df_solar.loc['area_FPV','Input'])/1000000) + (pyo.value(model.x2)*df_wind.loc['required_area_turbine','Input'])

print("\n Total Area System = %4.2f [km2]" %(total_area)) 

print("\n PV Installed Capacity = %4.2f [kWp]" %(installed_power_PV)) 
PV_CAPEX = installed_power_PV*pyo.value(model.PV_CAPEX)
print("\n PV CAPEX = %4.2f [EUR]" %(PV_CAPEX))
PV_OPEX = installed_power_PV*df_economic.loc['OPEX_solar_total','Input']
print("\n PV OPEX = %4.2f [EUR/year]" %(PV_OPEX))

print("\n Offshore Wind Installed Capacity = %4.2f [kW]" %(pyo.value(model.x2)*df_wind.loc['P_turbine','Input']))
OffshoreWind_CAPEX = pyo.value(model.x2)*pyo.value(model.W_CAPEX)
print("\n Offshore Wind CAPEX = %4.2f [EUR]" %(OffshoreWind_CAPEX))
OffshoreWind_OPEX = pyo.value(model.x2)*df_economic.loc['OPEX_wind','Input']
print("\n Offshore Wind OPEX = %4.2f [EUR/year]" %(OffshoreWind_OPEX))

Storage_CAPEX = pyo.value(model.x3)*pyo.value(model.CAPEX_Storage)
print("\n Storage CAPEX = %4.2f [EUR]" %(Storage_CAPEX))
Storage_OPEX = pyo.value(model.x3)*df_economic.loc['OPEX_battery','Input']
print("\n Storage OPEX = %4.2f [EUR/year]" %(Storage_OPEX))

Electrolyser_CAPEX = pyo.value(model.x4)*pyo.value(model.CAPEX_electrolyser)
print("\n Electrolyzer CAPEX = %4.2f [EUR]" %(Electrolyser_CAPEX))
Electrolyser_OPEX = pyo.value(model.x4)*df_economic.loc['OPEX_Electrolysis','Input']
print("\n Electrolyzer OPEX = %4.2f [EUR/year]" %(Electrolyser_OPEX))

Compressor_CAPEX = (pyo.value(model.x5)+pyo.value(model.x5b))*pyo.value(model.CAPEX_compressor)
print("\n Compressor CAPEX = %4.2f [EUR]" %(Compressor_CAPEX))
Compressor_OPEX = (pyo.value(model.x5)+pyo.value(model.x5b))*df_economic.loc['OPEX_compressor','Input']
print("\n Compressor OPEX = %4.2f [EUR/year]" %(Compressor_OPEX))

H2_Storage_CAPEX = pyo.value(model.x6)*df_economic.loc['CAPEX_H2_storage','Input']
print("\n H2 Storage CAPEX = %4.2f [EUR]" %(H2_Storage_CAPEX))
H2_Storage_OPEX = pyo.value(model.x6)*df_economic.loc['OPEX_H2_storage','Input']
print("\n H2 Storage OPEX = %4.2f [EUR/year]" %(H2_Storage_OPEX))


max_flow = max(pyo.value(model.H2_Export_h[i]) for i in model.T)

#Saving Results in lists to Postprocessing, export to DF, create graphs, etc...:
EPV = []                #Total PV production 
EW = []                 #Total Wind Production 
AEP_h = []              #Total electricity production per hour
E_curtailed = []        #Total Curtailed
E_used = []             #Total Used
EUsed_Grid_h = []
EUsed_Battery_h = []
EUsed_H2_h = []
EUsed_Grid_H2 = []      #for filling graphs  
AE_Export = []
SoC = [] 
E_Export_H2=[]          #for filling graphs               
E_Total_Consumed=[]     #for filling graphs            
E_H2_Total = []         #total electricity used for H2 to be accounted in OPEX!
H2_flow = []            #Total h2 production 
H2_EL_PIPE = []
H2_STO_PIPE = []
H2_Export = []   
SoC_H2 = []
# power_variation = []
           
for i in model.T:      
        EPV.append(pyo.value(model.EPV_h[i]))
        EW.append(pyo.value(model.EW_h[i]))  
        AEP_h.append(EW[i-1] + EPV[i-1])                            #Total Electricity Produced
        E_curtailed.append(pyo.value(model.ECurtailed_h[i]))
        E_used.append(pyo.value(model.EUsed_h[i]))                  #Total Electricity Used
        EUsed_Grid_h.append(pyo.value(model.EUsed_Grid_h[i]))       #Used to Grid Directly
        EUsed_Battery_h.append(pyo.value(model.EUsed_Battery_h[i])) #Used to Baterry Directly
        EUsed_H2_h.append(pyo.value(model.EUsed_H2_h[i]))           #Used to H2 Directly
        EUsed_Grid_H2.append(EUsed_Grid_h[i-1] + EUsed_H2_h[i-1])   #Used Grid and H2 (for filling graphs) 
        AE_Export.append(pyo.value(model.EBattery_Grid_h[i])+pyo.value(model.EUsed_Grid_h[i]))        
        E_Export_H2.append(EUsed_H2_h[i-1]+AE_Export[i-1])          #(for filling graphs) electricity to grid direct and from battery plus electricity direct to hydrogen production
        E_H2_Total.append(pyo.value(model.E_H2_Total_h[i]))            
        E_Total_Consumed.append(E_H2_Total[i-1]+AE_Export[i-1])     #(for filling graphs) electricity direct to grid and H2 production + electricity from battery to to grid + CMP + EL         
        H2_flow.append(pyo.value(model.H2_flow_h[i]))
        H2_Export.append(pyo.value(model.H2_Export_h[i]))
        H2_EL_PIPE.append(pyo.value(model.H2_EL_PIPE_h[i]))
        H2_STO_PIPE.append(pyo.value(model.H2_STO_PIPE_h[i]))
        SoC.append(pyo.value(model.SoC_h[i])) 
        SoC_H2.append(pyo.value(model.h2Storage[i]))       
        # power_variation.append(pyo.value(model.variation[i]))   


# Totals coud be find using also: Variable = sum(pyo.value(model.EPV_h[i])+ pyo.value(model.EW_h[i]) for i in model.T)
#% Curtailed
print('\n Percentage Curtailed = %4.2f' %(sum(E_curtailed)/sum(AEP_h)*100),'%')

#LCOE
AEused =  sum(E_used)
# print('\n AEP [kWh] = ',AEused)
electricity_yearly_production = costs_production_flow(AEused, int(df_general.loc['system_lifetime','Input'])) #yearly production! 
np_AEP = np_calculator(electricity_yearly_production, df_general.loc['discount_rate','Input'])

LCOE_OPEX = PV_OPEX + OffshoreWind_OPEX + Storage_OPEX #+ Electrolyser_OPEX + Compressor_OPEX
LCOE_CAPEX = PV_CAPEX + OffshoreWind_CAPEX + Storage_CAPEX #+ Electrolyser_CAPEX + Compressor_CAPEX
LCOE_cost_flow = costs_production_flow(LCOE_OPEX, int(df_general.loc['system_lifetime','Input']))
LCOE_cost_flow[0] = LCOE_CAPEX
np_LCOE_total_cost = np_calculator(LCOE_cost_flow, df_general.loc['discount_rate','Input'])
 
LCOE = np_LCOE_total_cost/(np_AEP/1000) #EUR/MWh
print('\n LCoE = %4.2f EUR/MWh' %(LCOE))

#LCOH
AH2P = sum(H2_flow) #kg
h2_yearly_production = costs_production_flow(AH2P, int(df_general.loc['system_lifetime','Input'])) #yearly production! 
np_production_H2 = np_calculator(h2_yearly_production, df_general.loc['discount_rate','Input'])

                                                                                                                                                                 
LCOH_OPEX =  Electrolyser_OPEX + Compressor_OPEX  + H2_Storage_OPEX + df_Hydrogen.loc['OPEX_pipe','Input']  + (sum(pyo.value(model.E_H2_Total_h[i])*(LCOE/1000) for i in model.T))                                    
LCOH_CAPEX = Electrolyser_CAPEX + Compressor_CAPEX + H2_Storage_CAPEX + df_Hydrogen.loc['CAPEX_pipe','Input'] 

LCOH_cost_flow_H2 = costs_production_flow(LCOH_OPEX, int(df_general.loc['system_lifetime','Input']))
LCOH_cost_flow_H2[0] = LCOH_CAPEX
#Adding Investment costs of the stack over the lifetime
i=1
while df_Hydrogen.loc['lifetime_stack','Input']*i < df_general.loc['system_lifetime','Input']:
    LCOH_cost_flow_H2[df_Hydrogen.loc['lifetime_stack','Input']*i] = LCOH_cost_flow_H2[df_Hydrogen.loc['lifetime_stack','Input']*i] + (df_Hydrogen.loc['CAPEX_stack','Input']*pyo.value(model.x4))
    i = i + 1    
np_LCOH_total_cost_H2 = np_calculator(LCOH_cost_flow_H2, df_general.loc['discount_rate','Input'])

LCOH = np_LCOH_total_cost_H2/np_production_H2

print('\n LCoH = %4.2f EUR/kg' %(LCOH))

#Cashflows
total_revenue_year = sum(pyo.value(model.H2_Export_h[i])*pyo.value(model.H2_price) for i in model.T) + sum(pyo.value(model.Electricity_export[i])*df_electricity_prices.loc[i-1,'Average_2015_2022_[EUR/kWh]'] for i in model.T)
revenue_flow = costs_production_flow(total_revenue_year, int(df_general.loc['system_lifetime','Input']))

total_opex_year = LCOE_OPEX + LCOH_OPEX - (sum(pyo.value(model.E_H2_Total_h[i])*(LCOE/1000) for i in model.T))  
total_costs_flow = costs_production_flow(total_opex_year, int(df_general.loc['system_lifetime','Input']))
capex_total = LCOE_CAPEX +  LCOH_CAPEX
total_costs_flow[0]=capex_total
i=1
while df_Hydrogen.loc['lifetime_stack','Input']*i < df_general.loc['system_lifetime','Input']:
    total_costs_flow[df_Hydrogen.loc['lifetime_stack','Input']*i] = total_costs_flow[df_Hydrogen.loc['lifetime_stack','Input']*i] + (df_Hydrogen.loc['CAPEX_stack','Input']*pyo.value(model.x4))
    i = i + 1   

net_cashflow = revenue_flow - total_costs_flow
IRR = npf.irr(net_cashflow)
NPV = np_calculator(net_cashflow, df_general.loc['discount_rate','Input'])

print('\n NPV = %4.2f' %(NPV))
print('\n IRR = %4.2f' %(IRR))

print('\n NPV Check (Objective Function = %4.2f' %(model.ObjFunction()))

# NPV could also be calculated using Numpy-Financial
# NPV =npf.npv(df_general.loc['discount_rate','Input'], net_cashflow)
# print('\n NPV2 = %4.2f' %(NPV))


#*********PLOT CURVES***********
x = [t for t in model.T]          # time varialbe for the plots 0 to 8759
font_size = 20

#Power Sources
plt.figure(figsize=(12,6))
plt.plot(x[0:8759], EPV[0:8759], label='PV', color='#f1c232') 
plt.plot(x[0:8759], EW[0:8759], label='Wind', color='#2f3c68') 
plt.title("Total Production EWind & EPV", fontsize=font_size)
plt.xlabel("Hours [h]", fontsize=font_size)
plt.ylabel("Power [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.legend()
plt.grid(True)
plt.show()

#SoC H2
plt.figure(figsize=(12,6))
plt.plot(x[0:8759],  SoC_H2[0:8759],'r--', label='Hydrogen Storage') 
plt.title("Hydrogen Storage Level", fontsize=font_size)
plt.xlabel("Hours [h]", fontsize=font_size)
plt.ylabel("Hydrogen Storage [kg]", fontsize=font_size)
plt.grid(True)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.legend()
plt.show()


#Time frame analysed
tf_start =  0
tf_end = 8759

#Production Electricity Distribution 
fig, (ax1) = plt.subplots(figsize=(12,6))
ax1.plot(x[tf_start:tf_end], AEP_h[tf_start:tf_end],'k', label='_nolegend_') 
ax1.fill_between(x[tf_start:tf_end],EUsed_Grid_h[tf_start:tf_end], color = '#012d7f', label='Export Cable')
ax1.fill_between(x[tf_start:tf_end],EUsed_Grid_H2[tf_start:tf_end], EUsed_Grid_h[tf_start:tf_end], color = '#38761d', label='Hydrogen Production')
ax1.fill_between(x[tf_start:tf_end],E_used[tf_start:tf_end],EUsed_Grid_H2[tf_start:tf_end], color = 'gold', label='Battery')
ax1.fill_between(x[tf_start:tf_end], AEP_h[tf_start:tf_end],E_used[tf_start:tf_end], color = 'red', label='Curtailed')
ax1.set_title("Produced Electricity Distribution", fontsize=font_size)
ax1.set_xlabel("Hours [h]", fontsize=font_size)
ax1.set_ylabel("Energy [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()

#Electricity Distribution Export/Consumption (Grid, Compressor, Hydrogen)
fig, (ax2) = plt.subplots(figsize=(12,6))
ax2.plot(x[tf_start:tf_end], E_Total_Consumed[tf_start:tf_end],'k', label='Total Electricity Used') 
ax2.fill_between(x[tf_start:tf_end],EUsed_Grid_h[tf_start:tf_end], color = '#012d7f', alpha=0.2, label='Export Cable')
ax2.fill_between(x[tf_start:tf_end],AE_Export[tf_start:tf_end], EUsed_Grid_h[tf_start:tf_end], color = '#012d7f', label='From Battery to Cable')
ax2.fill_between(x[tf_start:tf_end],E_Export_H2[tf_start:tf_end], EUsed_Grid_h[tf_start:tf_end], color = '#38761d',alpha=0.2, label='Consumed Hydrogen')
ax2.fill_between(x[tf_start:tf_end],E_Total_Consumed[tf_start:tf_end], E_Export_H2[tf_start:tf_end], color = '#38761d', label='From Battery to Hydrogen')
ax2.set_title("Used Electricity Distribution", fontsize=font_size)
ax2.set_xlabel("Hours [h]", fontsize=font_size)
ax2.set_ylabel("Energy [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()

#Hydrogen Production
fig, (ax3) = plt.subplots(figsize=(12,6))
ax3.plot(x[tf_start:tf_end],  H2_flow[tf_start:tf_end],color = '#38761d',linestyle='dashed', alpha=0.3, label='Total Hydrogen Produced') 
ax3.fill_between(x[tf_start:tf_end],H2_EL_PIPE[tf_start:tf_end], color = '#38761d', alpha=0.2, label='from Electrolyser')
ax3.fill_between(x[tf_start:tf_end],H2_Export[tf_start:tf_end], H2_EL_PIPE[tf_start:tf_end] ,color = '#38761d', label='from Storage')
ax3.set_title("H2 Production and Exported", fontsize=font_size)
ax3.set_xlabel("Hours [h]", fontsize=font_size)
ax3.set_ylabel("Hydrogen [kg]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()



#SoC
fig, (ax4) = plt.subplots(figsize=(12,6))
ax4.plot(x[tf_start:tf_end],  SoC[tf_start:tf_end],'r--', label='State of Charge') 
ax4.legend()
ax4.set_title("State of Charge", fontsize=font_size)
ax4.set_xlabel("Hours [h]", fontsize=font_size)
ax4.set_ylabel("SoC [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.tight_layout()
plt.grid(True)
plt.legend()


#Time frame analysed
tf_start =  3500
tf_end = 4000


#Power Sources
plt.figure(figsize=(12,6))
plt.plot(x[tf_start:tf_end], EPV[tf_start:tf_end], label='PV', color='#f1c232') 
plt.plot(x[tf_start:tf_end], EW[tf_start:tf_end], label='Wind', color='#2f3c68') 
plt.title("Total Production EWind & EPV", fontsize=font_size)
plt.xlabel("Hours [h]", fontsize=font_size)
plt.ylabel("Power [1e6, kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.legend()
plt.grid(True)
plt.show()

#Production Electricity Distribution 
fig, (ax5) = plt.subplots(figsize=(12,6))
ax5.plot(x[tf_start:tf_end], AEP_h[tf_start:tf_end],'k', label='_nolegend_') 
ax5.fill_between(x[tf_start:tf_end],EUsed_Grid_h[tf_start:tf_end], color = '#012d7f', label='Export Cable')
ax5.fill_between(x[tf_start:tf_end],EUsed_Grid_H2[tf_start:tf_end], EUsed_Grid_h[tf_start:tf_end], color = '#38761d', label='Hydrogen Production')
ax5.fill_between(x[tf_start:tf_end],E_used[tf_start:tf_end],EUsed_Grid_H2[tf_start:tf_end], color = 'gold', label='Battery')
ax5.fill_between(x[tf_start:tf_end], AEP_h[tf_start:tf_end],E_used[tf_start:tf_end], color = 'red', label='Curtailed')
ax5.set_title("Produced Electricity Distribution", fontsize=font_size)
ax5.set_xlabel("Hours [h]", fontsize=font_size)
ax5.set_ylabel("Energy [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()

#Electricity Distribution Export/Consumption (Grid, Compressor, Hydrogen)
fig, (ax6) = plt.subplots(figsize=(12,6))
ax6.plot(x[tf_start:tf_end], E_Total_Consumed[tf_start:tf_end],'k', label='Total Electricity Used') 
ax6.fill_between(x[tf_start:tf_end],EUsed_Grid_h[tf_start:tf_end], color = '#012d7f', alpha=0.2, label='Export Cable')
ax6.fill_between(x[tf_start:tf_end],AE_Export[tf_start:tf_end], EUsed_Grid_h[tf_start:tf_end], color = '#012d7f', label='From Battery to Cable')
ax6.fill_between(x[tf_start:tf_end],E_Export_H2[tf_start:tf_end], EUsed_Grid_h[tf_start:tf_end], color = '#38761d',alpha=0.2, label='Consumed Hydrogen')
ax6.fill_between(x[tf_start:tf_end],E_Total_Consumed[tf_start:tf_end], E_Export_H2[tf_start:tf_end], color = '#38761d', label='From Battery to Hydrogen')
ax6.set_title("Used Electricity Distribution", fontsize=font_size)
ax6.set_xlabel("Hours [h]", fontsize=font_size)
ax6.set_ylabel("Energy [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()

#Hydrogen Production
fig, (ax7) = plt.subplots(figsize=(12,6))
ax7.plot(x[tf_start:tf_end],  H2_flow[tf_start:tf_end],color = '#38761d',linestyle='dashed', alpha=0.3, label='Total Hydrogen Produced') 
ax7.fill_between(x[tf_start:tf_end],H2_EL_PIPE[tf_start:tf_end], color = '#38761d', alpha=0.2, label='from Electrolyser')
ax7.fill_between(x[tf_start:tf_end],H2_Export[tf_start:tf_end], H2_EL_PIPE[tf_start:tf_end] ,color = '#38761d', label='from Storage')
ax7.set_title("H2 Production and Exported", fontsize=font_size)
ax7.set_xlabel("Hours [h]", fontsize=font_size)
ax7.set_ylabel("Hydrogen [kg]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()

#SoC
fig, (ax8) = plt.subplots(figsize=(12,6))
ax8.plot(x[tf_start:tf_end],  SoC[tf_start:tf_end],'r--', label='State of Charge') 
ax8.legend()
ax8.set_title("State of Charge", fontsize=font_size)
ax8.set_xlabel("Hours [h]", fontsize=font_size)
ax8.set_ylabel("SoC [kWh]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.tight_layout()
plt.grid(True)
plt.legend()

#SoC H2
plt.figure(figsize=(12,6))
plt.plot(x[tf_start:tf_end],  SoC_H2[tf_start:tf_end],'r--', label='Hydrogen Storage') 
plt.title("Hydrogen Storage Level", fontsize=font_size)
plt.xlabel("Hours [h]", fontsize=font_size)
plt.ylabel("Hydrogen Storage [kg]", fontsize=font_size)
plt.xticks(fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.grid(True)
plt.legend()
plt.show()

# #Variation
# fig, (ax5) = plt.subplots(figsize=(12,6))
# ax5.plot(x[tf_start:tf_end],  power_variation[tf_start:tf_end],'r--', label='Variation Electricity Electrolyser') 
# ax5.legend()
# ax5.set_title("Variation", fontsize=font_size)
# ax5.set_xlabel("Hours [h]", fontsize=font_size)
# ax5.set_ylabel("Variation [kWh]", fontsize=font_size)
# plt.tight_layout()
# plt.grid(True)


   
# Sensitivity Analysis  
Sensitivity={}
PV_Sensitivity = {}
h2_price=[2,3,4,5,6,7,8,9,10,11,12,13,14,15]
for x in h2_price:
    model.H2_price=x             # Making the Parameter Mutuable=True, now I change the variable to perform a sensitivity analysis. 
    results = opt.solve(model) 
    Sensitivity[x,'Installed Capacity Electrolyser [MW]']= pyo.value(model.x4)/1000
    PV_Sensitivity[x,'FPV Area [km2]'] = (pyo.value(model.x1)*df_solar.loc['FPV_kWpm2','Input']*df_solar.loc['area_FPV','Input'])/1000000
print(Sensitivity)    

#Creating Graph:
Y=[Sensitivity[x,'Installed Capacity Electrolyser [MW]'] for x in h2_price]
X=[x for x in h2_price]
plt.plot(X,Y)
plt.ylabel('Installed Capacity [MW]')
plt.xlabel('Hydrogen Price [EUR/kg]')   
plt.grid(True)
plt.legend()
plt.xticks([2,3,4,5,6,7,8,9,10,11,12,13,14,15])
plt.show()

#Creating Graph:
Y=[PV_Sensitivity[x,'FPV Area [km2]'] for x in h2_price]
X=[x for x in h2_price]
plt.plot(X,Y)
plt.ylabel('FPV Area [km2]')
plt.xlabel('Hydrogen Price [EUR/kg]')   
plt.grid(True)
plt.legend()
plt.xticks([2,3,4,5,6,7,8,9,10,11,12,13,14,15])
plt.show()

# Sensitivity={}
# electricity_price=[0.055,0.060,0.065,0.070,0.075,0.080,0.085,0.090,0.095,0.100,0.105,0.110,0.115,0.120,0.125,0.13]
# for x in electricity_price:
#     model.electricity_price=x             # Making the Parameter Mutuable=True, now I change the variable to perform a sensitivity analysis. 
#     results = opt.solve(model) 
#     AEP = sum(pyo.value(model.EPV_h[i]) + pyo.value(model.EW_h[i]) for i in model.T) 
#     curtailed = sum(pyo.value(model.ECurtailed_h[i]) for i in model.T)
#     Sensitivity[x,'% Curtailed']= (curtailed/AEP)*100

# print(Sensitivity)    

# #Creating Graph:
# Y=[Sensitivity[x,'% Curtailed'] for x in electricity_price]
# X=[x for x in electricity_price]
# plt.plot(X,Y)
# plt.ylabel('% Curtailed')
# plt.xlabel('Electricity Price [EUR/kWh]')   
# plt.grid(True)
# plt.legend()
# plt.xticks([0.05, 0.06, 0.07, 0.08, 0.09, 0.1, 0.11, 0.12, 0.13])
# plt.show()



# # Sensitivity Analysis Hybrid Energy System Hydrogen Price
# Sensitivity={}
# h2_price=[4,4.2,4.4,4.6,4.7,4.8,4.9,5,5.2,5.4,5.6,5.8,6,6.2,6.4,6.6,6.8,7,7.2,7.4,7.5]
# for x in h2_price:
#     model.H2_price=x             # Making the Parameter Mutuable=True, now I change the variable to perform a sensitivity analysis. 
#     results = opt.solve(model) 
#     Sensitivity[x,'Electrolyser Installed Power']=  pyo.value(model.x4)/1000

# print(Sensitivity)    

# #Creating Graph:
# Y=[Sensitivity[x,'Electrolyser Installed Power'] for x in h2_price]
# X=[x for x in h2_price]
# plt.plot(X,Y)
# plt.ylabel('Electrolyser Installed Power [MW]')
# plt.xlabel('Hydrogen Price [EUR/kg]')   
# plt.grid(True)
# plt.legend()
# plt.xticks([4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5])
# plt.show()

# # Sensitivity Analysis  Hybrid Energy System Electricity Price
# Sensitivity={}
# electricity_price=[0.055,0.060,0.065,0.070,0.075,0.080,0.085,0.090,0.095,0.100,0.105,0.110,0.115,0.120,0.125,0.13]
# for x in electricity_price:
#     model.electricity_price=x             # Making the Parameter Mutuable=True, now I change the variable to perform a sensitivity analysis. 
#     results = opt.solve(model) 
#     Sensitivity[x,'Electrolyser Installed Power']= pyo.value(model.x4)/1000

# print(Sensitivity)    

# #Creating Graph:
# Y=[Sensitivity[x,'Electrolyser Installed Power'] for x in electricity_price]
# X=[x for x in electricity_price]
# plt.plot(X,Y)
# plt.ylabel('Electrolyser Capacity [MW]', fontsize=14)
# plt.xlabel('Electricity Price [EUR/kWh]', fontsize=14)   
# plt.grid(True)
# plt.legend()
# plt.show()



#Export to excel

# df_c1 = pd.DataFrame(c1)   
# df_SoC = pd.DataFrame(SoC)     
# df_Battery_charge = pd.DataFrame(Battery_charge)      
# df_Battery_discharge = pd.DataFrame(Battery_discharge)  
# df_PV_production = pd.DataFrame(PV_production) 
# df_EPV_used = pd.DataFrame(EPV_used) 
# df_PVcurtailment = pd.DataFrame(PVcurtailment) 


# with pd.ExcelWriter(' Python output.xlsx') as writer:
#     df_c1[0].to_excel(writer, sheet_name="Balance Constraint")
#     df_SoC[0].to_excel(writer, sheet_name="SoC")
#     df_PV_production[0].to_excel(writer, sheet_name="PV Production")
#     df_Battery_charge[0].to_excel(writer, sheet_name="Charge")
#     df_Battery_discharge[0].to_excel(writer, sheet_name="Discharge")
#     df_EPV_used[0].to_excel(writer, sheet_name="PV")
#     df_PVcurtailment[0].to_excel(writer, sheet_name="PV Curtailed")