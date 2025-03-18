from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
from textblob import TextBlob
from fastapi.responses import StreamingResponse

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
        df = pd.read_excel(BytesIO(contents), engine='openpyxl')

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
        df.insert(1, "Sentiment", sentiments)
        df.insert(2, "Sentiment Score", scores)

        # Save updated file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
            writer.book.close()

        output.seek(0)

        # Return file as download
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=sentiment_feedback.xlsx"}
        )

    except Exception as e:
        return {"error": str(e)}
