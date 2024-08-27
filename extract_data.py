import pandas as pd
import os 

def extract_data_by_asset_type(folder_path):
    """
    Returns:
        list: A list of dictionaries, each with 'Asset Type' as the key.
    """
    all_structured_data_list = []

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(file_path)
            data_list = df.to_dict(orient='records')

            structured_data_list = [] # Set 'Asset Type' as the key for each row
            for row in data_list:
                asset_type = row.pop('Asset Type')  # Remove 'Asset Type' and store its value
                structured_row = {asset_type: row}  # Use the 'Asset Type' value as the key
                structured_data_list.append(structured_row)

            all_structured_data_list.extend(structured_data_list)

    return all_structured_data_list

folder_path = 'Batch-input-Page\Templates'
result = extract_data_by_asset_type(folder_path)
print(result)