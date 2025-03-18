from fastapi import FastAPI, File, UploadFile, Response
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
from textblob import TextBlob
import openpyxl  # Ensure openpyxl is installed

app = FastAPI()

# ✅ Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Change to specific domains in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Root endpoint for health check
@app.get("/")
def root():
    return {"message": "FastAPI Server is Running!"}

# ✅ Excel Upload & Sentiment Analysis Endpoint
@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    try:
        # ✅ Read the uploaded Excel file
        contents = await file.read()
        input_stream = BytesIO(contents)

        # ✅ Load Excel into DataFrame
        df = pd.read_excel(input_stream, engine="openpyxl")

        # ✅ Ensure the first column exists
        if df.shape[1] < 1:
            return {"error": "The Excel file does not contain enough columns."}

        # ✅ Rename the first column as "Feedback"
        df.rename(columns={df.columns[0]: "Feedback"}, inplace=True)

        # ✅ Check for empty values & fill missing feedback
        df["Feedback"].fillna("", inplace=True)

        # ✅ Sentiment Analysis
        df["Sentiment"] = df["Feedback"].apply(lambda x: (
            "Positive" if TextBlob(str(x)).sentiment.polarity > 0 
            else "Negative" if TextBlob(str(x)).sentiment.polarity < 0 
            else "Neutral"
        ))
        df["Sentiment Score"] = df["Feedback"].apply(lambda x: TextBlob(str(x)).sentiment.polarity)

        # ✅ Save the processed file
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sentiment Analysis")
            writer.close()  # ✅ Ensure all data is written

        output.seek(0)  # ✅ Reset stream position

        # ✅ Return processed file
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
