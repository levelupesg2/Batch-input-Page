from flask import Flask, request, jsonify

import tempfile
from extract_data import process_asset_data, extract_data_by_asset_type
from werkzeug.utils import secure_filename
import requests

app = Flask(__name__)

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Get the JSON data from the request
        data = request.get_json()
        if not isinstance(data, list):
            return jsonify({"error": "Invalid input format. Expected a list of JSON objects."}), 400
        
        # Prepare a list to store the results
        results = []

        # Process each JSON object in the list
        for item in data:
            file_link = item.get('file_link')
            company_id = item.get('companyId')

            if not file_link or not company_id:
                return jsonify({"error": "Missing 'file_link' or 'companyId' in one of the JSON objects."}), 400

            # Check if the file_link is a local path
            if file_link.startswith("temp\\"):
                temp_file_path = file_link
            else:
                # Download the file using the file_link
                response = requests.get(file_link)
                if response.status_code != 200:
                    return jsonify({"error": f"Failed to download the file from {file_link}"}), 400

                # Save the file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                    temp_file.write(response.content)
                    temp_file_path = temp_file.name

            # Extract data from the file
            data_extracted = extract_data_by_asset_type(temp_file_path)
            if not data_extracted:
                return jsonify({"error": f"No data found in the file {file_link}"}), 400

            # Call the functions based on asset type
            result = process_asset_data(data_extracted)

            # Append the result to the results list
            results.append({
                "companyId": company_id,
                "file_link": file_link,
                "result": result
            })

        # Return the list of results
        return jsonify(results), 200
    
    except Exception as e:
        return jsonify({"error": f"Failed to process the request: {str(e)}"}), 400
