import pandas as pd
import json
from Calculations import (Refrigerants, Heat_and_Steam, Other_Stationary, Purchased_Electricity, Company_Vehicles, Natural_Gas_func)

def extract_data_by_asset_type(file_path):
    """
    Extracts data by asset type from a single Excel file, skipping the first two rows.
    Returns:
        DataFrame: A DataFrame with the processed data.
        list: A list of column names.
    """
    
    df = pd.read_excel(file_path, skiprows=[1, 2]) # skip the first two rows
    df.columns = df.columns.str.strip().str.lower()
    

    if 'asset type' not in df.columns:
        print("Warning: 'asset type' column not found in the Excel file.")
        return pd.DataFrame(), []
    
    # df = df.dropna(subset=['asset type'])
    
    return df

def process_asset_data(df):
    """
    Calls functions from Calculations.py based on the 'Asset Type' in the input DataFrame and returns the results as JSON.

    """
    processed_results = []

    for index, row in df.iterrows():
        asset_type = row['asset type'].strip().lower() 

        if asset_type == 'refrigerants':
            result = Refrigerants(
                Actual_estimate=row['actual/estimated'],
                reporting_year=row['reporting year'],
                Equipment_type=row['equipment type'],
                Purpose_stage=row['purpose stage'],
                Refrigerant_type=row['refrigerant type'],
                Refrigerant_lost_kg=row['refrigerant recovered (kg)'],
                method=row['method'],
                total_refrigerant_charge=row['total refrigerant charge (kg)']
            )
            processed_results.append(result)

        elif asset_type == 'heat and steam':
            result = Heat_and_Steam(
                Actual_estimate=row['actual/estimated'],
                Reporting_Year=row['reporting year'],
                Typology=row['typology'],
                value_type=row['value type'],
                consumtion=row['consumption (kwh)'],
                Total_spend=row['total spend'],
                currency_Type=row['currency']
            )
            processed_results.append(result)

        elif asset_type == 'other stationary':
            result = Other_Stationary(
                Actual_estimate=row['actual/estimated'],
                Reporting_Year=row['reporting year'],
                Fuel_type=row['fuel type'],
                Fuel_Unit=row['fuel unit'],
                value_type=row['value type'],
                Consumption=row['consumption'],
                Total_spend=row['total spend'],
                currency=row['currency']
            )
            processed_results.append(result)

        elif asset_type == 'purchased electricity':
            result = Purchased_Electricity(
                Actual_estimate=row['actual/estimated'],
                Country=row['country'],
                Tariff=row['tariff'],
                Reporting_Year=row['reporting year'],
                value_type=row['value type'],
                Consumption_kWh=row['consumption (kwh)'],
                currency=row['currency'],
                Total_spend=row['total spend'],
                Coal=row['coal'],
                Natural_Gas=row['natural gas'],
                Nuclear=row['nuclear'],
                Renewables=row['renewables'],
                Other_Fuel=row['other fuel'],
                Coal_percent=row['coal percent'],
                Natural_Gas_percent=row['natural gas percent'],
                Nuclear_percent=row['nuclear percent'],
                Renewables_percent=row['renewables percent'],
                Other_Fuel_percent=row['other fuel percent']
            )
            processed_results.append(result)

        elif asset_type == 'natural gas':
            result = Natural_Gas_func(
                Actual_estimate=row['actual/estimated'],
                reporting_year=row['reporting year'],
                Meter_Read_Units=row['meter read units'],
                value_type=row['value type'],
                Consumption=row['consumption'],
                Total_spend=row['total spend'],
                currency=row['currency']
            )
            processed_results.append(result)

        elif asset_type in ['company vehicles (distance-based)', 'company vehicles (fuel-based)']:
            if 'distance based' in row:  
                result = Company_Vehicles(
                    Actual_estimate=row['actual/estimated'],
                    Activity_Type=row['activity type'],
                    Reporting_Year=row['reporting year'],
                    Method='distance based',
                    Vehicle_category=row.get('vehicle category', ""),
                    Vehicle_Type=row.get('vehicle type', ""),
                    Fuel_type_Laden=row.get('fuel type/laden', "Diesel"),
                    Unit_distance_travelled=row.get('unit distance travelled', "miles"),
                    Distance_travelled=row.get('distance travelled', 0)
                )
            else:  
                result = Company_Vehicles(
                    Actual_estimate=row['actual/estimated'],
                    Activity_Type=None,
                    Reporting_Year=row['reporting year'],
                    Method='fuel based',
                    Fuel_type=row.get('fuel type', "Aviation spirit"),
                    Fuel_Amount_in_litres=row.get('fuel_amount_in_litres', 0)
                )
            processed_results.append(result)

        else:
            return(f"Warning: No processing function found for asset type '{asset_type}'")

    return json.dumps(processed_results, indent=2, default=str)



# file_path = 'Batch-input-Page/Templates/Purchased_Electricity.xlsx'
# file_path = 'Batch-input-Page/Templates/Natural_Gas.xlsx'
# file_path = 'Batch-input-Page/Templates/Heat_and_Steam.xlsx'
# file_path = 'Batch-input-Page/Templates/Refrigerants.xlsx'
# file_path = 'Batch-input-Page/Templates/Other_Stationary.xlsx'
# file_path = 'Batch-input-Page/Templates/Company_Vehicles_Fuel_based.xlsx'
file_path = 'Batch-input-Page/Templates/Company_Vehicles_Distance_based.xlsx'

df = extract_data_by_asset_type(file_path)
if not df.empty:
    processed_data_json = process_asset_data(df)
    print(processed_data_json)
else:
    print("No data to process.")