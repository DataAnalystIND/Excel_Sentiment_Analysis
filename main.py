from fastapi import FastAPI, UploadFile, File
import pandas as pd
from io import BytesIO
from fastapi.responses import Response
import uvicorn

app = FastAPI()

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    try:
        # Read the uploaded Excel file
        df = pd.read_excel(file.file, engine="openpyxl")

        # Ensure column exists
        if "Feedback" not in df.columns:
            return {"error": "Column 'Feedback' not found in the uploaded file"}

        # Perform simple sentiment analysis
        df["Sentiment"] = df["Feedback"].apply(lambda x: "Positive" if "good" in str(x).lower() else "Negative")

        # Save processed file to memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        # Ensure data is written before sending response
        output.seek(0)

        return Response(
            output.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=sentiment_feedback.xlsx"},
        )

    except Exception as e:
        return {"error": str(e)}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=10000)
