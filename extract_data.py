import pandas as pd
import json
from Calculations import (Refrigerants, Heat_and_Steam, Other_Stationary, Purchased_Electricity, Company_Vehicles, Natural_Gas_func)

def extract_data_by_asset_type(file_path):
    """
    Extracts data by asset type from a single Excel file, skipping the first two rows and keeping the third row as headers.

    Returns:
        list: A list of dictionaries, each containing 'Asset Type' as the key and corresponding data.
    """
    all_structured_data_list = []
    df = pd.read_excel(file_path, skiprows=[1,2])  # skipping first 2 rows 
    # print("Column names in the Excel file:", df.columns.tolist())
    df.columns = df.columns.str.strip().str.lower()  
    data_list = df.to_dict(orient='records')

    structured_data_list = [] 
    for row in data_list:

        asset_type = row.pop('asset type') 
        
        if asset_type is None:
            print(f"Warning: 'asset type' not found in row: {row}")
            continue

        structured_row = {asset_type: row} 
        structured_data_list.append(structured_row)

    all_structured_data_list.extend(structured_data_list)

    return all_structured_data_list , df.columns.tolist()


def purchased_electricity(Actual_estimate, Country , Tariff, Reporting_Year, value_type, 
                          Reporting_periods_list=[2023], EF_years_list=[2023],
                          Consumption_kWh=0, currency="", Total_spend=0, 
                          Coal=0, Natural_Gas=0, Nuclear=0, Renewables=0, Other_Fuel=0,
                          Coal_percent=0, Natural_Gas_percent=0, Nuclear_percent=0, 
                          Renewables_percent=0, Other_Fuel_percent=0):
    return (f"Processed Purchased Electricity data: Country={Country}, Tariff={Tariff}, "
            f"Year={Reporting_Year}, Value Type={value_type}, Consumption={Consumption_kWh} kWh, "
            f"Total Spend={Total_spend} {currency}, Coal={Coal} ({Coal_percent}%), "
            f"Natural Gas={Natural_Gas} ({Natural_Gas_percent}%), Nuclear={Nuclear} ({Nuclear_percent}%), "
            f"Renewables={Renewables} ({Renewables_percent}%), Other Fuel={Other_Fuel} ({Other_Fuel_percent}%), "
            f"Reporting Periods={Reporting_periods_list}, EF Years={EF_years_list}")


def company_vehicles(data):
    return f"Processed data for Company Vehicles: {data}"



asset_type_functions = {
    'Refrigerants': Refrigerants,
    'Heat_and_Steam': Heat_and_Steam,
    'Other_Stationary': Other_Stationary,
    'Purchased_Electricity': purchased_electricity, ########## change this### to Purchased_Electricity
    'Company_Vehicles': company_vehicles, ########## change this###
    'Natural_Gas_func': Natural_Gas_func
}
def process_asset_data(structured_data_list):
    """
    Calls functions from Calculations.py based on the 'Asset Type' in the input data and returns the results as JSON.

    Args:
        data (list): A list of dictionaries containing asset data.

    Returns:
        str: A JSON string containing the results of the calculations.
    """
    processed_results = []

    for data in structured_data_list:
        for asset_type, asset_data in data.items():
            if asset_type == 'Refrigerants':
                result = Refrigerants(
                    Actual_estimate = None ,
                    reporting_year  =asset_data.get('reporting year'),
                     Equipment_type =asset_data.get('equipment type'),
                     Purpose_stage =asset_data.get('purpose stage'),
                     Refrigerant_type =asset_data.get('refrigerant type'),
                     Refrigerant_lost_kg =asset_data.get('refrigerant recovered (kg)'),
                    method=asset_data.get('method'),
                    total_refrigerant_charge=asset_data.get('total refrigerant charge (kg)')
                )
                processed_results.append({asset_type: result})
            elif asset_type == 'Heat and Steam':
                result = Heat_and_Steam(
                    Actual_estimate = None ,
                    Reporting_Year=asset_data.get('reporting year'),
                    Typology=asset_data.get('typology'),
                    value_type=asset_data.get('value type'),
                    consumtion=asset_data.get('consumption (kWh)'),
                    Total_spend=asset_data.get('total spend'),
                    currency_Type=asset_data.get('currency')
                )
                processed_results.append({asset_type: result})
            elif asset_type == 'Other Stationary':
                result = Other_Stationary(
                    Actual_estimate = None ,
                    Reporting_Year=asset_data.get('reporting year'),
                    Fuel_type=asset_data.get('fuel type'),
                    Fuel_Unit=asset_data.get('fuel unit'),
                    value_type=asset_data.get('value type'),
                    Consumption=asset_data.get('consumption'),
                    Total_spend=asset_data.get('total spend'),
                    currency=asset_data.get('currency')
                )
                processed_results.append({asset_type: result})             
            elif asset_type == 'Purchased Electricity':
                ########## change this### to Purchased_Electricity
                result = purchased_electricity( 
                    Actual_estimate = None ,
                    Country=asset_data.get('country'),
                    Tariff=asset_data.get('tariff'),
                    Reporting_Year=asset_data.get('reporting year'),
                    value_type=asset_data.get('value type'),
                    Consumption_kWh=asset_data.get('consumption (kWh)'),
                    currency=asset_data.get('currency'),
                    Total_spend=asset_data.get('total spend'),
                    Coal=asset_data.get('coal'),
                    Natural_Gas=asset_data.get('natural gas'),
                    Nuclear=asset_data.get('nuclear'),
                    Renewables=asset_data.get('renewables'),
                    Other_Fuel=asset_data.get('other fuel'),
                    Coal_percent=asset_data.get('coal percent'),
                    Natural_Gas_percent=asset_data.get('natural gas percent'),
                    Nuclear_percent=asset_data.get('nuclear percent'),
                    Renewables_percent=asset_data.get('renewables percent'),
                    Other_Fuel_percent=asset_data.get('other fuel percent')
                )
                processed_results.append({asset_type: result})                
            elif asset_type == 'Natural Gas':
                result = Natural_Gas_func(
                    Actual_estimate = None ,
                    reporting_year=asset_data.get('reporting year'),
                    Meter_Read_Units=asset_data.get('meter read units'),
                    value_type=asset_data.get('value type'),
                    Consumption=asset_data.get('consumption'),
                    Total_spend=asset_data.get('total spend'),
                    currency=asset_data.get('currency')
                )
                processed_results.append({asset_type: result})
            ########## change this### ADD  Company_Vehicles 2 functions 
            else:
                process_function = asset_type_functions.get(asset_type)
                if process_function:
                    result = process_function(asset_data)
                    processed_results.append({asset_type: result})
                else:
                    print(f"Warning: No processing function found for asset type '{asset_type}'")

    return json.dumps(processed_results, indent=4, default=str)


file_path = 'Batch-input-Page/Templates/Refrigerants.xlsx'
result , headers = extract_data_by_asset_type(file_path)
# print(result[0],headers)
# print(headers)
# print(result[0])

json_result = process_asset_data(result)
print(json_result)