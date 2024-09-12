import pandas as pd
import numpy as np
from datetime import datetime
from extract_data import extract_data_by_asset_type

class EnergyDataQualityChecker:
    def __init__(self):
        self.current_year = datetime.now().year

    def load_data(self, file_path):
        return pd.read_excel(file_path, skiprows=[1, 2])

    def temporal_correlation(self, year):
        if year == self.current_year:
            return 1  # Excellent
        elif year == self.current_year - 1:
            return 2  # Very Good
        elif year == self.current_year - 2:
            return 3  # Good
        elif year == self.current_year - 3:
            return 4  # Poor
        else:
            return 5  # Very Poor

    def completeness(self, row, required_fields):
        filled_fields = sum(pd.notna(row[field]) for field in required_fields)
        completeness_score = filled_fields / len(required_fields)
        if completeness_score == 1:
            return 1  # Excellent
        elif completeness_score >= 0.8:
            return 2  # Very Good
        elif completeness_score >= 0.6:
            return 3  # Good
        elif completeness_score >= 0.4:
            return 4  # Poor
        else:
            return 5  # Very Poor

    def reliability(self, row):
        if row['Actual/Estimated'].lower() == 'actual':
            return 1  # Excellent
        elif row['Actual/Estimated'].lower() == 'estimated' and pd.notna(row['Assumption basis']):
            return 3  # Good
        else:
            return 5  # Very Poor

    def assess_natural_gas(self, data):
        required_fields = ['Asset Name', 'Asset Type', 'Reporting Year', 'Value Type', 
                           'Consumption', 'Meter Read Units', 'Total Spend', 'Currency', 
                           'Actual/Estimated', 'Evidence']
        data['Temporal Correlation'] = data['Reporting Year'].apply(self.temporal_correlation)
        data['Completeness'] = data.apply(lambda row: self.completeness(row, required_fields), axis=1)
        data['Reliability'] = data.apply(self.reliability, axis=1)
        return data

    def assess_purchased_electricity(self, data):
        required_fields = ['Asset Name', 'Asset Type', 'Country', 'Reporting Year', 'Tariff', 
                           'Value Type', 'Consumption (kWh)', 'Total Spend', 'Currency', 
                           'Coal', 'Natural Gas', 'Nuclear', 'Renewables', 'Other Fuel', 
                           'Actual/Estimated', 'Evidence']
        data['Temporal Correlation'] = data['Reporting Year'].apply(self.temporal_correlation)
        data['Completeness'] = data.apply(lambda row: self.completeness(row, required_fields), axis=1)
        data['Reliability'] = data.apply(self.reliability, axis=1)
        return data

    def assess_company_vehicles(self, data):
        required_fields = ['Asset Name', 'Asset Type', 'Reporting Year', 'Method', 
                           'Fuel Type', 'Fuel Amount in litres', 'Actual/Estimated', 'Evidence']
        data['Temporal Correlation'] = data['Reporting Year'].apply(self.temporal_correlation)
        data['Completeness'] = data.apply(lambda row: self.completeness(row, required_fields), axis=1)
        data['Reliability'] = data.apply(self.reliability, axis=1)
        return data

    def assess_other_stationary(self, data):
        required_fields = ['Asset Type', 'Reporting Year', 'Fuel Type', 'Fuel Unit', 
                           'Value Type', 'Consumption', 'Total Spend', 'Currency', 
                           'Actual/Estimated', 'Evidence']
        data['Temporal Correlation'] = data['Reporting Year'].apply(self.temporal_correlation)
        data['Completeness'] = data.apply(lambda row: self.completeness(row, required_fields), axis=1)
        data['Reliability'] = data.apply(self.reliability, axis=1)
        return data

    def assess_refrigerants(self, data):
        required_fields = ['Asset Name', 'Asset Type', 'Reporting Year', 'Method', 
                           'Equipment Type', 'Purpose Stage', 'Refrigerant Type', 
                           'Refrigerant Recovered (kg)', 'Total Refrigerant Charge (kg)', 
                           'Actual/Estimated', 'Evidence']
        data['Temporal Correlation'] = data['Reporting Year'].apply(self.temporal_correlation)
        data['Completeness'] = data.apply(lambda row: self.completeness(row, required_fields), axis=1)
        data['Reliability'] = data.apply(self.reliability, axis=1)
        return data

    def assess_heat_and_steam(self, data):
        required_fields = ['Asset Name', 'Asset Type', 'Reporting Year', 'Value Type', 
                           'Consumption (kWh)', 'Typology', 'Total Spend', 'Currency', 
                           'Actual/Estimated', 'Evidence']
        data['Temporal Correlation'] = data['Reporting Year'].apply(self.temporal_correlation)
        data['Completeness'] = data.apply(lambda row: self.completeness(row, required_fields), axis=1)
        data['Reliability'] = data.apply(self.reliability, axis=1)
        return data

    def calculate_dqi(self, data):
        data['DQI'] = (data['Temporal Correlation'] + data['Completeness'] + data['Reliability']) / 3
        data['DQ Problem?'] = data['DQI'].apply(lambda x: 'Yes' if x > 3 else 'No')
        return data

    def assess_data_quality(self, file_path, data_type):
        data = self.load_data(file_path)
        if data_type == 'natural_gas':
            data = self.assess_natural_gas(data)
        elif data_type == 'purchased_electricity':
            data = self.assess_purchased_electricity(data)
        elif data_type == 'company_vehicles':
            data = self.assess_company_vehicles(data)
        elif data_type == 'other_stationary':
            data = self.assess_other_stationary(data)
        elif data_type == 'refrigerants':
            data = self.assess_refrigerants(data)
        elif data_type == 'heat_and_steam':
            data = self.assess_heat_and_steam(data)
        else:
            raise ValueError(f"Unknown data type: {data_type}")
        
        return self.calculate_dqi(data)

# Usage example:
checker = EnergyDataQualityChecker()
natural_gas_dq = checker.assess_data_quality('./Templates/Natural_Gas.xlsx', 'natural_gas')
print(natural_gas_dq)
# natural_gas_dq.to_excel('./natural_gas_dq_assessment.xlsx', index=False)