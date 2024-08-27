import pandas as pd
# import json
# from Calculations import (Refrigerants, Heat_and_Steam, Other_Stationary, Purchased_Electricity, Company_Vehicles, Natural_Gas_func)

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

file_path = 'Batch-input-Page\Templates\Refrigerants.xlsx'
result = extract_data_by_asset_type(file_path)
print(result[0])
