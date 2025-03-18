from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
from textblob import TextBlob

app = FastAPI()

# ✅ Fixing CORS Issue
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
        # Read the uploaded Excel file
        contents = await file.read()
        df = pd.read_excel(BytesIO(contents))

        # Check if Column A exists
        if df.shape[1] < 1:
            return {"error": "Column A (Feedback) is missing"}

        # Rename Column A to 'Feedback'
        df.rename(columns={df.columns[0]: "Feedback"}, inplace=True)

        # Perform Sentiment Analysis
        sentiments = []
        scores = []
        for feedback in df["Feedback"]:
            if pd.isna(feedback):  # Handle empty cells
                sentiments.append("Neutral")
                scores.append(0)
            else:
                sentiment_score = TextBlob(str(feedback)).sentiment.polarity
                sentiments.append("Positive" if sentiment_score > 0 else "Negative" if sentiment_score < 0 else "Neutral")
                scores.append(sentiment_score)

        # Add Sentiment Results to DataFrame
        df["Sentiment"] = sentiments
        df["Sentiment Score"] = scores

        # Save updated file
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return {
            "filename": "updated_feedback.xlsx",
            "message": "File processed successfully",
            "data": df.to_dict(orient="records"),
        }
    except Exception as e:
        return {"error": str(e)}

