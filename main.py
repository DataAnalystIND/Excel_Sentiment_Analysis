from fastapi import FastAPI, File, UploadFile, Response
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
from textblob import TextBlob
import openpyxl

app = FastAPI()

# ✅ Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Root endpoint to support both GET & HEAD
@app.api_route("/", methods=["GET", "HEAD"])
def root():
    return {"message": "FastAPI Server is Running!"}

# ✅ Upload Excel and Analyze Sentiment
@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        input_stream = BytesIO(contents)
        df = pd.read_excel(input_stream, engine="openpyxl")

        # ✅ Check if there is at least one column
        if df.shape[1] < 1:
            return {"error": "The Excel file does not contain enough columns."}

        df.rename(columns={df.columns[0]: "Feedback"}, inplace=True)
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
            writer.close()

        output.seek(0)

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
