import pandas as pd
import os
from flask import Flask, request, send_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    if file.filename == "":
        return "No selected file", 400

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    try:
        # Read the Excel file correctly
        df = pd.read_excel(file_path, engine="openpyxl")

        # Process (example: add a column)
        df["Processed_Column"] = "Processed"

        # Save with explicit format to prevent corruption
        processed_path = os.path.join(PROCESSED_FOLDER, "processed_" + file.filename)
        df.to_excel(processed_path, index=False, engine="openpyxl")

        # Ensure Windows doesn't block the file
        final_path = processed_path.replace(".xlsx", "_safe.xlsx")
        os.rename(processed_path, final_path)

        return send_file(final_path, as_attachment=True)

    except Exception as e:
        return f"Error processing file: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)

