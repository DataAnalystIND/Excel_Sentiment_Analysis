from fastapi import FastAPI, File, UploadFile, Response
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
from textblob import TextBlob
import openpyxl  # Ensure openpyxl is installed

app = FastAPI()

# ✅ CORS Fix
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Change to specific domains in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    try:
        # ✅ Read the uploaded Excel file
        contents = await file.read()
        input_stream = BytesIO(contents)
        df = pd.read_excel(input_stream, engine="openpyxl")

        # ✅ Ensure Column A (Feedback) exists
        if df.shape[1] < 1:
            return {"error": "Column A (Feedback) is missing"}

        # ✅ Rename the first column as "Feedback"
        df.rename(columns={df.columns[0]: "Feedback"}, inplace=True)

        # ✅ Sentiment Analysis Processing
        df["Sentiment"] = df["Feedback"].apply(lambda x: (
            "Positive" if TextBlob(str(x)).sentiment.polarity > 0 
            else "Negative" if TextBlob(str(x)).sentiment.polarity < 0 
            else "Neutral"
        ))
        df["Sentiment Score"] = df["Feedback"].apply(lambda x: TextBlob(str(x)).sentiment.polarity)

        # ✅ Save the updated file to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
            writer.close()  # ✅ Ensure data is properly flushed

        output.seek(0)  # ✅ Reset stream position

        # ✅ Return file as an HTTP response
        return Response(
            content=output.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": "attachment; filename=sentiment_feedback.xlsx",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        )

    except Exception as e:
        return {"error": str(e)}
