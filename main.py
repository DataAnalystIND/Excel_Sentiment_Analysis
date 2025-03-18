from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
import pandas as pd
import os
import shutil

app = FastAPI()

def remove_windows_block(file_path):
    """Remove Windows security block from the file"""
    temp_file = file_path + "_safe.xlsx"  # Create a temporary copy
    shutil.copy2(file_path, temp_file)  # Copy file to a new one
    os.remove(file_path)  # Delete the old file
    os.rename(temp_file, file_path)  # Rename temp file back to original

@app.post("/process/")
async def process_file(file: UploadFile = File(...)):
    try:
        # Read the uploaded Excel file
        df = pd.read_excel(file.file)
        
        # Example Processing: Convert text to uppercase in Column A
        if 'A' in df.columns:
            df['A'] = df['A'].astype(str).str.upper()
        
        # Save processed file
        output_file = "processed_feedback.xlsx"
        df.to_excel(output_file, index=False)
        
        # Remove Windows Security Block
        remove_windows_block(output_file)
        
        return FileResponse(output_file, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="processed_feedback.xlsx")
    
    except Exception as e:
        return {"error": str(e)}
