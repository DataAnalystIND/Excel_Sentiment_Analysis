from fastapi import FastAPI, File, UploadFile
import pandas as pd
from textblob import TextBlob
from io import BytesIO
from fastapi.responses import StreamingResponse

app = FastAPI()

@app.get("/")
def home():
    return {"message": "Sentiment Analysis API is Running!"}

@app.post("/upload-excel/")
async def process_excel(file: UploadFile = File(...)):
    try:
        # Read the uploaded file
        contents = await file.read()
        excel_data = pd.read_excel(BytesIO(contents))

        # Check if 'Feedback' column exists
        if 'Feedback' not in excel_data.columns:
            return {"error": "Column A must have 'Feedback' as the header."}

        # Perform Sentiment Analysis
        sentiments = []
        scores = []
        
        for feedback in excel_data['Feedback']:
            analysis = TextBlob(str(feedback))
            sentiment_score = analysis.sentiment.polarity  # Score between -1 to 1
            sentiments.append("Positive" if sentiment_score > 0 else "Negative")
            scores.append(sentiment_score)

        # Add results to new columns
        excel_data["Sentiment"] = sentiments
        excel_data["Sentiment Score"] = scores

        # Save updated file
        output = BytesIO()
        excel_data.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=sentiment_feedback.xlsx"})

    except Exception as e:
        return {"error": str(e)}
