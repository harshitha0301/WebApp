from flask import Flask, render_template, request
import pandas as pd
import os
import openpyxl
import tempfile
from applicationinsights import TelemetryClient


ExcelFileRead = Flask(__name__, template_folder="templates")

INSTRUMENTATION_KEY = "b77836d0-17fc-446a-b283-f0d24bf171db"
tc = TelemetryClient(INSTRUMENTATION_KEY)

def validate_excel(filepath):
    df = pd.read_excel(filepath)

    # Data type validation
    expected_data_types = {
        'Artist': str,      
        'Song': str
    }

    # Remove rows with invalid data types
    for column, expected_type in expected_data_types.items():
        if df[column].dtype != expected_type:
            df = df[df[column].apply(lambda x: isinstance(x, expected_type))]

    # Removing duplicates
    df = df.drop_duplicates()

    # Removing rows with missing values
    df = df.dropna()
    return df

@ExcelFileRead.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            if "file" not in request.files:
                return "No file part"
            
            file = request.files["file"]
            if file.filename == "":
                return "No selected file"
            
            if file:
                with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                    temp_path = temp_file.name
                    file.save(temp_path)
                    
                    valid_data = validate_excel(temp_path)
                    tc.track_event("File Uploaded", { "filename": file.filename })
                    return render_template("result.html", valid_data=valid_data.to_html(index=False))
                
        except Exception as e:
            tc.track_exception()
            return "Error: " + str(e)
        
    return render_template("index.html")

if __name__ == "__main__":
    ExcelFileRead.run(debug=True)

