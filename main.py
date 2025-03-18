from fastapi import FastAPI, UploadFile, File
import pandas as pd
from io import BytesIO
from fastapi.responses import Response
import uvicorn

app = FastAPI()

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    try:
        # Read uploaded file
        df = pd.read_excel(file.file, engine="openpyxl")

        # Validate if the required column exists
        if "Feedback" not in df.columns:
            return {"error": "Column 'Feedback' not found in the uploaded file"}

        # Apply simple sentiment analysis
        df["Sentiment"] = df["Feedback"].astype(str).apply(lambda x: "Positive" if "good" in x.lower() else "Negative")

        # Save processed file to memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl", mode="xlsx") as writer:
            df.to_excel(writer, index=False, sheet_name="Sentiment Analysis")

        # Ensure file is properly written
        output.seek(0)

        return Response(
            output.read(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=sentiment_feedback.xlsx"},
        )

    except Exception as e:
        return {"error": str(e)}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=10000)
