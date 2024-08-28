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

    return all_structured_data_list

def call_functions_based_on_asset_type(data):
    """
    Calls specific functions from Calculations.py based on the 'Asset Type' in the input data and returns the results as JSON.

    Args:
        data (list): A list of dictionaries containing asset data.

    Returns:
        str: A JSON string containing the results of the calculations.
    """
    results_json = []

    asset_function_map = {
        'refrigerants': Refrigerants,
        'heat_and_steam': Heat_and_Steam,
        'other_stationary': Other_Stationary,
        'purchased_electricity': Purchased_Electricity,
        'company_vehicles': Company_Vehicles,
        'natural_gas': Natural_Gas_func
    }

    required_parameters = {
        'refrigerants': {
            'reporting_year': {'excel_col': 'reporting_year', 'default': None},
            'equipment_type': {'excel_col': 'equipment_type', 'default': None},
            'purpose_stage': {'excel_col': 'purpose_stage', 'default': None},
            'refrigerant_type': {'excel_col': 'refrigerant_type', 'default': None},
            'refrigerant_lost_kg': {'excel_col': 'refrigerant_recovered_(kg)', 'default': 0},
            'method': {'excel_col': 'method', 'default': 'Unknown Method'},
            'total_refrigerant_charge': {'excel_col': 'total_refrigerant_charge_(kg)', 'default': 0}
        },
        'heat_and_steam': {
            'reporting_year': {'excel_col': 'reporting_year', 'default': None},
            'typology': {'excel_col': 'typology', 'default': 'General'},
            'value_type': {'excel_col': 'value_type', 'default': 'Estimate'},
            'consumption': {'excel_col': 'consumption', 'default': 0},
            'total_spend': {'excel_col': 'total_spend', 'default': 0},
            'currency_type': {'excel_col': 'currency_type', 'default': 'USD'}
        },
        'other_stationary': {
            'reporting_year': {'excel_col': 'reporting_year', 'default': None},
            'fuel_type': {'excel_col': 'fuel_type', 'default': 'Unknown'},
            'fuel_unit': {'excel_col': 'fuel_unit', 'default': 'liters'},
            'value_type': {'excel_col': 'value_type', 'default': 'Estimate'},
            'consumption': {'excel_col': 'consumption', 'default': 0},
            'total_spend': {'excel_col': 'total_spend', 'default': 0},
            'currency': {'excel_col': 'currency', 'default': 'USD'}
        },
        'purchased_electricity': {
            'country': {'excel_col': 'country', 'default': 'Unknown'},
            'tariff': {'excel_col': 'tariff', 'default': 'Standard'},
            'reporting_year': {'excel_col': 'reporting_year', 'default': None},
            'value_type': {'excel_col': 'value_type', 'default': 'Estimate'},
            'consumption_kwh': {'excel_col': 'consumption_kwh', 'default': 0},
            'currency': {'excel_col': 'currency', 'default': 'USD'},
            'total_spend': {'excel_col': 'total_spend', 'default': 0},
            'coal': {'excel_col': 'coal', 'default': 0},
            'natural_gas': {'excel_col': 'natural_gas', 'default': 0},
            'nuclear': {'excel_col': 'nuclear', 'default': 0},
            'renewables': {'excel_col': 'renewables', 'default': 0},
            'other_fuel': {'excel_col': 'other_fuel', 'default': 0},
            'coal_percent': {'excel_col': 'coal_percent', 'default': 0.0},
            'natural_gas_percent': {'excel_col': 'natural_gas_percent', 'default': 0.0},
            'nuclear_percent': {'excel_col': 'nuclear_percent', 'default': 0.0},
            'renewables_percent': {'excel_col': 'renewables_percent', 'default': 0.0},
            'other_fuel_percent': {'excel_col': 'other_fuel_percent', 'default': 0.0}
        },
        'company_vehicles': {
            'activity_type': {'excel_col': 'activity_type', 'default': 'Transport'},
            'reporting_year': {'excel_col': 'reporting_year', 'default': None},
            'method': {'excel_col': 'method', 'default': 'Standard'},
            'vehicle_category': {'excel_col': 'vehicle_category', 'default': 'Light'},
            'vehicle_type': {'excel_col': 'vehicle_type', 'default': 'Car'},
            'fuel_type': {'excel_col': 'fuel_type', 'default': 'Petrol'},
            'fuel_amount_in_litres': {'excel_col': 'fuel_amount_in_litres', 'default': 0},
            'fuel_type_laden': {'excel_col': 'fuel_type_laden', 'default': 'None'},
            'unit_distance_travelled': {'excel_col': 'unit_distance_travelled', 'default': 'km'},
            'distance_travelled': {'excel_col': 'distance_travelled', 'default': 0}
        },
        'natural_gas': {
            'reporting_year': {'excel_col': 'reporting_year', 'default': None},
            'meter_read_units': {'excel_col': 'meter_read_units', 'default': 'cubic meters'},
            'value_type': {'excel_col': 'value_type', 'default': 'Estimate'},
            'consumption': {'excel_col': 'consumption', 'default': 0},
            'total_spend': {'excel_col': 'total_spend', 'default': 0},
            'currency': {'excel_col': 'currency', 'default': 'USD'}
        }
    }

    for item in data:
        for asset_type, details in item.items():
            normalized_asset_type = asset_type.strip().lower().replace(' ', '_')

            calculation_function = asset_function_map.get(normalized_asset_type)

            
            if calculation_function:
                # Extract only the parameters that the function requires, providing defaults if missing
                function_params = {}
                for param, info in required_parameters[normalized_asset_type].items():
                    
                    function_params[param] = details.get(info['excel_col'], info['default'])

                result = calculation_function(**function_params)
                print(result)
             

    return json.dumps(results_json)

file_path = 'Batch-input-Page\Templates\Refrigerants.xlsx'
result = extract_data_by_asset_type(file_path)
# print(result[0])
call_functions_based_on_asset_type([result[5]])
