import pandas as pd
import numpy as np
import math as mt

def conversion_to_USD(currency_Type, total_spend): 
    df_currency = pd.read_excel("Batch-input-Page/Data/conversion to USD.xlsx")
    df_currency[['2021','2022','2023']] = df_currency[['2021','2022','2023']].replace(["no data" , None] , 0 )
    df_currency[['2021','2022','2023']] = df_currency[['2021','2022','2023']].astype('float')
    filter_df = df_currency[df_currency['Currency from'].str.lower() == str(currency_Type).lower()]
    
    if not filter_df.empty:
        conversion_rate = round(filter_df['2023'].iloc[0], 3)
        total_spend_after_conversion = round(total_spend / conversion_rate, 3)
        return total_spend_after_conversion
    return "Currency are not correct"

def Refrigerants(Actual_estimate ,reporting_year , 
                   Equipment_type ,
                   Purpose_stage ,
                   Refrigerant_type ,
                   Refrigerant_lost_kg ,
                  method , Reporting_periods_list =[2023],
                  EF_years_list = [2023] ,total_refrigerant_charge = 0) : 
    
    # reporting period list = [ from 2021 to 2023 ] in my Excel 
    if Actual_estimate == None :
        Actual_estimate = "actual" 
    # Calculate Refrigerant_Lost_kg 
    Refrigerant_Lost_kg = 0
    Simplified_Material_Balance_method = "Simplified Material Balance method"
    if method == Simplified_Material_Balance_method.lower() : 
        Refrigerant_Lost_kg = Refrigerant_lost_kg

    #*********************************************************************
    # Calculate EF_Year 
    EF_years_lists = EF_years_list
    index = Reporting_periods_list.index(reporting_year)
    EF_Year= EF_years_lists[index]

    #*********************************************************************
    # calculate Emission factor (kgCO2e/kg)
    df = pd.read_excel("Batch-input-Page/Data/BEIS_sheet(S2).xlsx")

    # Define the column names
    columns = ['Level 3', 'Year', 'Level 5', 'GHG Conversion Factor']

    # Set the column names
    df = df[columns]
    df.columns
    # Filter the DataFrame to match the refrigerant type, EF year, and column text
    filtered_df = df[(df['Level 3'].str.lower() == str(Refrigerant_type).lower()) &
                    (df['Year'] == EF_Year ) &
                    (df['Level 5'] == 'Kyoto products')]

    # If a match is found, return the GHG conversion factor
    if not filtered_df.empty:
        Emission_factor_kgCO2e_kg = filtered_df['GHG Conversion Factor'].iloc[0]

    #*********************************************************************
   
    # calculate NON KYOTO Emission factor (kgCO2e/kg)2 
    filtered_df1 = df[(df['Level 3'].str.lower() == str(Refrigerant_type).lower()) &
                    (df['Year'] == EF_Year ) &
                    (df['Level 5'] == 'Non Kyoto')]
    NON_KYOTO_Emission_factor_kgCO2e_kg_2 = filtered_df1['GHG Conversion Factor'].iloc[0]

    #*********************************************************************
    #calculate Total Emissions kgCO2e 
    equipment_df = pd.read_excel("Batch-input-Page/Data/Refrigerant Equipment.xlsx")

    # Simplified Material Balance method
    Refrigerants_results = {}
    Simplified_Material_Balance_method = "Simplified Material Balance method"
    if method == Simplified_Material_Balance_method.lower() : 
        Total_Emissionskg_CO2e = Refrigerant_Lost_kg * Emission_factor_kgCO2e_kg

    # Other methods
    else:
        # Get the emission factor for the equipment type and purpose stage
        emission_factor_equipment = equipment_df.loc[equipment_df['Refrigerant Equipment'].str.lower() == str(Equipment_type).lower() , [col for col in equipment_df.columns ]].iloc[0].tolist()
        purpose_stage_index = [str(col).lower() for col in equipment_df.columns ].index(str(Purpose_stage).lower())
        emission_factor_equipment = emission_factor_equipment[purpose_stage_index]
        # Calculate emissions
        total_refrigerant_charge = total_refrigerant_charge
        Total_Emissionskg_CO2e = float(total_refrigerant_charge) * emission_factor_equipment * Emission_factor_kgCO2e_kg

    Refrigerants_Emissions_tCO2e = (Total_Emissionskg_CO2e) / 1000 
    Refrigerants_results = {
        "Refrigerant_Lost_kg": Refrigerant_Lost_kg,
        "EF_Year": EF_Year,
        "Emission_factor_kgCO2e_kg": Emission_factor_kgCO2e_kg,
        "NON_KYOTO_Emission_factor_kgCO2e_kg_2": NON_KYOTO_Emission_factor_kgCO2e_kg_2,
        "Total_Emissionskg_CO"
        "2e": Total_Emissionskg_CO2e,
        "Refrigerants_Emissions_tCO2e": Refrigerants_Emissions_tCO2e
    }

    return Refrigerants_results

# print(Refrigerants(2023 , "Condensing Units" , "Installation" ,"R410B" , 10 , "Simplified Material Balance method" , [2020,2021,2022,2023],[2019,2020,2021,2022],30))

def Heat_and_Steam(Actual_estimate , Reporting_Year,
                   Typology, 
                   value_type,
                   Reporting_periods_list=[2020,2021,2022,2023,2024],
                   EF_years_list=[2020,2021,2022,2023,2024] ,
                   consumtion = 0, 
                    Total_spend=0 ,
                    currency_Type= "" ):
    # Ensure Reporting_periods_list and EF_years_list are lists
    if Actual_estimate == None :
        Actual_estimate = "actual" 
    if not isinstance(Reporting_periods_list, list):
        Reporting_periods_list = [Reporting_periods_list]
    if not isinstance(EF_years_list, list):
        EF_years_list = [EF_years_list]

    # Calculate EF_Year
    index = Reporting_periods_list.index(Reporting_Year)
    EF_Year = EF_years_list[index]
    
    # Load data and perform calculations
    df = pd.read_excel("Batch-input-Page/Data/BEIS_sheet(S2).xlsx")

    # Define the column names
    columns = ['Level 2', 'Year', 'Level 3', "GHG/Unit", 'GHG Conversion Factor']

    # Set the column names
    df = df[columns]

    # Filter the DataFrame to match the refrigerant type, EF year, and column text
    filtered_df = df[(df['Level 2'] == "Heat and steam") &
                     (df['Year'] == EF_Year) &
                     (df['Level 3'].str.lower() == Typology.lower()) &
                     (df['GHG/Unit'] == 'kg CO2e')]

    Emission_factor_kgCO2e_USD = 0
    if not filtered_df.empty:
        Emission_factor_kgCO2e_USD = filtered_df['GHG Conversion Factor'].iloc[0]
        
    Total_Emissions_kgCO2e = 0
    Consumption_name  , total_spend_name= "Consumption" , "Total Spend"
    if str(value_type).lower() == Consumption_name.lower() : 
        if consumtion != 0 and Emission_factor_kgCO2e_USD != 0:
            Total_Emissions_kgCO2e = consumtion * Emission_factor_kgCO2e_USD

    elif str(value_type).lower() == total_spend_name.lower() :   
        total_spend_after_conversion = conversion_to_USD(currency_Type, Total_spend)
        if total_spend_after_conversion != 0 and Emission_factor_kgCO2e_USD != 0:
            Total_Emissions_kgCO2e = total_spend_after_conversion * Emission_factor_kgCO2e_USD
        
    Heat_and_Steam_Emissions_tCO2e = Total_Emissions_kgCO2e / 1000 if Total_Emissions_kgCO2e is not None else None
    
    Heat_and_Steam_results = {
        "EF_Year": EF_Year,
        "Emission_factor_kgCO2e_USD": Emission_factor_kgCO2e_USD,
        "Total_Emissions_kgCO2e": Total_Emissions_kgCO2e,
        "Heat_and_Steam_Emissions_tCO2e": Heat_and_Steam_Emissions_tCO2e
    }

    return Heat_and_Steam_results

# print(Heat_and_Steam(2023 ,"Onsite heat and steam","Consumption" , [2020 , 2021 ,2022 , 2023]  , [2019,2020,2021,2022] ,1500))

def Other_Stationary(Actual_estimate , Reporting_Year, Fuel_type, Fuel_Unit, value_type ,  
                       Reporting_periods_list=[2020,2021,2022,2023,2024],
                     EF_years_list=[2020,2021,2022,2023,2024], Consumption = 0 ,Total_spend = 0 , currency = ""):

    if Actual_estimate == None :
        Actual_estimate = "actual"
    if not isinstance(Reporting_periods_list, list):
        Reporting_periods_list = [Reporting_periods_list]
    if not isinstance(EF_years_list, list):
        EF_years_list = [EF_years_list]

    # Calculate EF_Year
    # print(f"Reporting_periods_list : {Reporting_periods_list}")
    index = Reporting_periods_list.index(Reporting_Year)
    EF_Year = EF_years_list[index]

    # Load data and perform calculations
    df = pd.read_excel("Batch-input-Page/Data/BEIS_sheet(S2).xlsx")

    # Filter the DataFrame to match the refrigerant type, EF year, and column text
    filtered_df = df[(df['Level 3'].str.lower() == str(Fuel_type).lower()) &
                     (df['UOM'].str.lower() == str(Fuel_Unit).lower()) ]

    filtered_df = filtered_df[
                     (df['Year'] == int(EF_Year)) &
                     (df['GHG/Unit'] == "kg CO2e")
                     ]

    # If a match is found, return the GHG conversion factor
    Emission_factor_kgCO2e_consumption_unit = 0
    if not filtered_df.empty:
        Emission_factor_kgCO2e_consumption_unit = filtered_df['GHG Conversion Factor'].iloc[0]

    Total_Emissions_kgCO2e = 0
    
    Consumption_name , total_spend_name= "Consumption" , "Total Spend"
    
    # print(f"Consumption :{Consumption}  ||  Emission_factor_kgCO2e_consumption_unit : {Emission_factor_kgCO2e_consumption_unit}")
    
    if str(value_type).lower() == Consumption_name.lower() : 
        if Consumption is not None and Emission_factor_kgCO2e_consumption_unit is not None:
             Total_Emissions_kgCO2e = float(Consumption) * Emission_factor_kgCO2e_consumption_unit
             
    elif str(value_type).lower() == total_spend_name.lower() :     
        total_spend_after_conversion = conversion_to_USD(currency, Total_spend)
        if total_spend_after_conversion != 0 and Emission_factor_kgCO2e_consumption_unit != 0:
            Total_Emissions_kgCO2e = total_spend_after_conversion * Emission_factor_kgCO2e_consumption_unit
            
    Other_Fuel_Emissions_tCO2e = (float(Total_Emissions_kgCO2e) / 1000) if Total_Emissions_kgCO2e is not None else None

    Other_Stationary = {
        "EF_Year": EF_Year,
        "Emission_factor_kgCO2e_kg": Emission_factor_kgCO2e_consumption_unit,
        "Total_Emissions_kgCO2e": Total_Emissions_kgCO2e,
        "Other_Fuel_Emissions_tCO2e": Other_Fuel_Emissions_tCO2e
    }

    return Other_Stationary
     
# print(Other_Stationary( 2022 ,"Butane" ,"kWh (Gross CV)" ,"Consumption" , [2021 ,2022 , 2023] , [2020,2021,2022]  , 1521 ))

def Purchased_Electricity(Actual_estimate , Country, Tariff, Reporting_Year, value_type, 
                          Reporting_periods_list=[2020,2021,2022,2023,2024], EF_years_list=[2020,2021,2022,2023,2024],
                           Consumption_kWh=0, currency="", Total_spend=0, 
                          Coal=0, Natural_Gas=0, Nuclear=0, Renewables=0, Other_Fuel=0,
                          Coal_percent=0, Natural_Gas_percent=0, Nuclear_percent=0, 
                          Renewables_percent=0, Other_Fuel_percent=0 ):
    if Actual_estimate == None :
        Actual_estimate = "actual"
    # Ensure Reporting_periods_list and EF_years_list are lists
    if not isinstance(Reporting_periods_list, list):
        Reporting_periods_list = [Reporting_periods_list]
    if not isinstance(EF_years_list, list):
        EF_years_list = [EF_years_list]

    # Calculate EF_Year
    index = Reporting_periods_list.index(Reporting_Year)
    EF_Year = EF_years_list[index]

    # Load data and perform calculations
    df = pd.read_excel("Batch-input-Page/Data/IEA_sheet(S2).xlsx")

    # Filter the DataFrame to match the country and EF year
    filtered_df = df[(df['Country/Region'].str.lower() == str(Country).lower()) &
                     (df['Year'] == EF_Year)]

    Location_Based_Emission_Factor_kgCO2e_kWh = 0
    if not filtered_df.empty:
        Location_Based_Emission_Factor_kgCO2e_kWh = filtered_df["Scope 2 only (kgCO2e/kWh)"].iloc[0]

    Location_Based_emissions_kgCO2e = 0
    Consumption_name , total_spend_name= "Consumption" , "Total Spend"
    
    # print(f"Consumption :{Consumption}  ||  Emission_factor_kgCO2e_consumption_unit : {Emission_factor_kgCO2e_consumption_unit}")
    
    if str(value_type).lower() == Consumption_name.lower() : 
        if Consumption_kWh != 0 and Location_Based_Emission_Factor_kgCO2e_kWh != 0:
            Location_Based_emissions_kgCO2e = Consumption_kWh * Location_Based_Emission_Factor_kgCO2e_kWh
            
    elif str(value_type).lower() == total_spend_name.lower():
        total_spend_after_conversion = conversion_to_USD(currency, Total_spend)
        if isinstance(total_spend_after_conversion, (int, float)) and Location_Based_Emission_Factor_kgCO2e_kWh != 0:
            Location_Based_emissions_kgCO2e = total_spend_after_conversion * Location_Based_Emission_Factor_kgCO2e_kWh    
        else:
            return f"Error: {total_spend_after_conversion}"

    # Calculate Market-Based Emission Factor (kgCO2e/kWh)
    ownsuppliermixtb_df = pd.read_excel("Batch-input-Page/Data/Own supplier mix.xlsx")
    electricity_marketb_ef_df = pd.read_excel("Batch-input-Page/Data/Own supplier EF.xlsx")

    ownsuppliermixtb_new_row = {
        "Site": "-", "Countries": str(Country).lower(), "EF year": EF_Year, 
        "Coal %": Coal_percent, "Natural Gas %": Natural_Gas_percent,
        "Nuclear %": Nuclear_percent, "Renewables %": Renewables_percent,
        "Other Fuel %": Other_Fuel_percent
    }
    electricity_marketb_ef_new_row = {
        "EF": "Emission factor (KgCO2/KWh)", "EF year": EF_Year, 
        "Coal": Coal, "Natural Gas": Natural_Gas, "Nuclear": Nuclear, 
        "Renewables": Renewables, "Other Fuel": Other_Fuel
    }

    ownsuppliermixtb_df = ownsuppliermixtb_df.append(ownsuppliermixtb_new_row, ignore_index=True)#################
    electricity_marketb_ef_df = electricity_marketb_ef_df.append(electricity_marketb_ef_new_row, ignore_index=True)###################

    Market_Based_Emission_Factor_kgCO2e_kWh = 0
    Grid_average = "Grid average"
    Renewable_procurement = "Renewable procurement"
    Own_supplier_mix = "Own supplier mix"
    if str(Tariff).lower() == Grid_average.lower():
        Market_Based_Emission_Factor_kgCO2e_kWh = Location_Based_Emission_Factor_kgCO2e_kWh
    elif str(Tariff).lower() == Renewable_procurement.lower():
        Market_Based_Emission_Factor_kgCO2e_kWh = 0
    elif str(Tariff).lower() == Own_supplier_mix.lower():
        fuel_columns = ['Coal %', 'Natural Gas %', 'Nuclear %', 'Renewables %', 'Other Fuel %']
        mix_row = ownsuppliermixtb_df.loc[
            (ownsuppliermixtb_df['Countries'].str.lower() == str(Country).lower()) & 
            (ownsuppliermixtb_df['EF year'] == EF_Year), fuel_columns
        ].values.flatten()

        ef_row = electricity_marketb_ef_df.loc[
            electricity_marketb_ef_df['EF year'] == EF_Year, 
            ['Coal', 'Natural Gas', 'Nuclear', 'Renewables', 'Other Fuel']
        ].values.flatten()
        
        # ef_row = ef_row[:len(mix_row)]################################
        ef_row = np.nan_to_num(ef_row)
        Fuel_mix_emissions_KgCO2_KWh = np.sum(mix_row * ef_row)

        Market_Based_Emission_Factor_kgCO2e_kWh = Fuel_mix_emissions_KgCO2_KWh if Fuel_mix_emissions_KgCO2_KWh != 0 else 0

    Market_Based_Emissions_kgCO2e = 0
    Consumption_name = "Consumption"
    Total_Spend_name = "Total Spend"
    if str(value_type).lower() == Consumption_name.lower():
        if Consumption_kWh != 0 and Market_Based_Emission_Factor_kgCO2e_kWh != 0:
            Market_Based_Emissions_kgCO2e = Consumption_kWh * Market_Based_Emission_Factor_kgCO2e_kWh
    elif str(value_type).lower() == Total_Spend_name.lower():
        total_spend_after_conversion = conversion_to_USD(currency.lower(), Total_spend)
        if isinstance(total_spend_after_conversion, (int, float)) and Market_Based_Emission_Factor_kgCO2e_kWh != 0:
            Market_Based_Emissions_kgCO2e = total_spend_after_conversion * Market_Based_Emission_Factor_kgCO2e_kWh
        else:
            return f"Error: {total_spend_after_conversion}"

    Purchased_Electricity_results = {
        "EF_Year": EF_Year,
        "Location_Based_Emission_Factor_kgCO2e_kWh": Location_Based_Emission_Factor_kgCO2e_kWh,
        "Location_Based_emissions_kgCO2e": Location_Based_emissions_kgCO2e,
        "Market_Based_Emission_Factor_kgCO2e_kWh": Market_Based_Emission_Factor_kgCO2e_kWh,
        "Market_Based_Emissions_kgCO2e": Market_Based_Emissions_kgCO2e
    }

    return Purchased_Electricity_results
# print(Purchased_Electricity(Country = "Egypt" ,Tariff = "Grid average" , 2023,"Consumption" ,[2021,2022,2023] ,[2020,2021 ,2022] ,1351.626 ))
# print(Purchased_Electricity("Egypt" ,"Renewable procurement" , 2023,"Consumption" ,[2021,2022,2023] ,[2020,2021 ,2022] ,1351.626 ))
# print(Purchased_Electricity("Egypt" ,"Own supplier mix" , 2023,
#                             "Total Spend" ,[2021,2022,2023] ,[2020,2021 ,2022] ,0,"ALL" ,1351.188,1200,2000,2051 ,3021,1533,0.25,0.124,0.32,0.48,0.46))



def Company_Vehicles(Actual_estimate ,Activity_Type,
                     Reporting_Year,
                     Method,
                     Vehicle_category = "",
                     Vehicle_Type = "", 
                     Fuel_type="Aviation spirit",
                     Fuel_Amount_in_litres=10, 
                     Fuel_type_Laden="Diesel", 
                     Unit_distance_travelled="miles",
                     Distance_travelled=5000, 
                     Reporting_periods_list=[2020,2021,2022,2023,2024],
                     EF_years_list=[2020,2021,2022,2023,2024]):
    if Actual_estimate == None :
        Actual_estimate = "actual"
    # Calculate EF_Year
    index = Reporting_periods_list.index(Reporting_Year)
    print(f"EF years : {EF_years_list} , Reporitn year list : {Reporting_periods_list}")
    EF_Year = EF_years_list[index]
    

    # Calculate Emission factor (kgCO2e/consumption unit)
    beisEFtb = pd.read_excel("Batch-input-Page/Data/BEIS_sheet(S2).xlsx")
    Emission_factor_kgCO2e_consumption_unit = 0

    if str(Activity_Type).lower() == "farm related" :
        Vehicle_category=""
        Vehicle_Type = ""

    # Perform the lookups based on the method
    if str(Method).lower() == "fuel based":
        condition1 = (
            (beisEFtb['UOM'] == "litres") &
            (beisEFtb['GHG/Unit'] == "kg CO2e") &
            (beisEFtb['Scope'] == "Scope 1") &
            (beisEFtb['Level 3'].str.lower() == Fuel_type.lower()) &
            (beisEFtb['Year'] == EF_Year)
        )
        result1 = beisEFtb.loc[condition1, 'GHG Conversion Factor']
        
        if not result1.empty:
            Emission_factor_kgCO2e_consumption_unit = result1.iloc[0]
    
    elif str(Method).lower() == "distance based":
        
        condition2 =beisEFtb [
            (beisEFtb['Year'] == EF_Year ) &
            (beisEFtb['Level 2'].str.lower() == (Vehicle_category).lower()) &
            (beisEFtb['Scope'] == "Scope 1") 
        ]
        # print(f"condition 2_1 : {condition2.shape}")
        
        condition2 = condition2[
            (beisEFtb['Level 3'].str.lower() == str(Vehicle_Type).lower()) &
            (beisEFtb['Level 5'].str.lower() == str(Fuel_type_Laden).lower()) &
            (beisEFtb['UOM'].str.lower() == str(Unit_distance_travelled).lower()) &
            (beisEFtb['GHG/Unit'] == "kg CO2e")
        ]
        
        # print(f"condition2_2 : {condition2.shape}")
        result2 = condition2.loc[: ,'GHG Conversion Factor']
        
        # print(f"result2 : {result2}")
        
        if not result2.empty:
            Emission_factor_kgCO2e_consumption_unit = result2.iloc[0]

      # Calculate total emissions
    Total_Emissions_kgCO2e = 0
    if str(Method).lower() == "fuel based" and Fuel_Amount_in_litres is not None and Emission_factor_kgCO2e_consumption_unit is not None:
        Total_Emissions_kgCO2e = Fuel_Amount_in_litres * Emission_factor_kgCO2e_consumption_unit
    elif str(Method).lower() == "distance based" and Distance_travelled is not None and Emission_factor_kgCO2e_consumption_unit is not None:
        Total_Emissions_kgCO2e = Distance_travelled * Emission_factor_kgCO2e_consumption_unit


    Company_vehicles_Emissions_tCO2e = Total_Emissions_kgCO2e / 1000

    FLAG_Emissions_tCO2e_or_no_flag = 0.0
    if str(Activity_Type).lower() in ["farm related", "non farm related"]:
        FLAG_Emissions_tCO2e_or_no_flag = float(Total_Emissions_kgCO2e) / 1000
        
    Company_Vehicles = {
        "EF_Year": EF_Year,
        "Emission_factor_kgCO2e_consumption_unit": Emission_factor_kgCO2e_consumption_unit,
        "Total_Emissions_kgCO2e": Total_Emissions_kgCO2e,
        "Company_vehicles_Emissions_tCO2e": Company_vehicles_Emissions_tCO2e,
        "FLAG_Emissions_tCO2e_or_no_flag": FLAG_Emissions_tCO2e_or_no_flag
    }
    
    return Company_Vehicles
 
# print(Company_Vehicles("Non-farm related" , 2023,"fuel based" ,None ,None,"", "LNG" ,135))
 
# print(Company_Vehicles("Non-farm related" , 2022 ,"distance based" ,"Gaseous fuels" ,"Butane",None,None,"Energy - Net CV","kWh (Net CV)",5031, [2020 ,2021 ,2022,2023,2024] , [2019,2020,2021,2022,2023])) 
 
 
def Natural_Gas_func(Actual_estimate ,reporting_year , Meter_Read_Units , value_type ,  
                     Consumption = 0 ,Total_spend = 0 , currency = "",
                   Reporting_periods_list=[2020,2021,2022,2023,2024] , 
                   EF_years_list=[2020,2021,2022,2023,2024]) :
    
    #*********************************************************************
    # Calculate EF_Year 
    if Actual_estimate == None :
        Actual_estimate = "actual"
    index = Reporting_periods_list.index(reporting_year)
    EF_Year= EF_years_list[index]
   
    #*********************************************************************
    # calculate Emission factor (kgCO2e/kg)
    df = pd.read_excel("Batch-input-Page/Data/BEIS_sheet(S2).xlsx")

    # Filter the DataFrame to match the refrigerant type, EF year, and column text
    filtered_df = df[
                (df['Level 3'] == 'Natural gas') &
                (df['UOM'].str.lower() == str(Meter_Read_Units).lower()) &
                (df['Year'] == EF_Year) &
                (df['Level 1'] == 'Fuels')
            ]

    Emission_factor_kgCO2e_consumption_unit=0
    # If a match is found, return the GHG conversion factor
    if not filtered_df.empty:
        Emission_factor_kgCO2e_consumption_unit = filtered_df['GHG Conversion Factor'].iloc[0]

    #*********************************************************************
    Total_Emissions_kgCO2e = 0
    Consumption_name = "Consumption"
    Total_Spend_name = "Total Spend"
    if str(value_type).lower() == Consumption_name.lower() :
        if Emission_factor_kgCO2e_consumption_unit != 0 and Consumption != 0 :
            Total_Emissions_kgCO2e =  Consumption * Emission_factor_kgCO2e_consumption_unit
            
    elif str(value_type).lower() == Total_Spend_name.lower() : 
        if Emission_factor_kgCO2e_consumption_unit != 0 and Total_spend != 0 :
            total_spend_after_conversion = conversion_to_USD(currency , Total_spend)
            Total_Emissions_kgCO2e =  total_spend_after_conversion * Emission_factor_kgCO2e_consumption_unit       
            
    Natural_Gas_results = {
        "EF_Year": EF_Year,
        "Emission_factor_kgCO2e_consumption_unit": Emission_factor_kgCO2e_consumption_unit,
        "Total_Emissions_kgCO2e": Total_Emissions_kgCO2e
    }

    return Natural_Gas_results


# print(Natural_Gas_func(2022,"kWh (Gross CV)" ,"consumption" ,1351,0,"",[2020,2021,2022,2023] ,[2019,2020,2021,2022]))
# print(Natural_Gas_func(2022,"kWh (Gross CV)" ,"total spend" ,0,135432.213,"ALL",[2020,2021,2022,2023] ,[2019,2020,2021,2022]))