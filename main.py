from fastapi import FastAPI, File, UploadFile, Response
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

        # Check if Column A (Feedback) exists
        if df.shape[1] < 1:
            return {"error": "Column A (Feedback) is missing"}

        # Rename Column A to 'Feedback'
        df.rename(columns={df.columns[0]: "Feedback"}, inplace=True)

        # Perform Sentiment Analysis
        df["Sentiment"] = df["Feedback"].apply(lambda x: (
            "Positive" if TextBlob(str(x)).sentiment.polarity > 0 
            else "Negative" if TextBlob(str(x)).sentiment.polarity < 0 
            else "Neutral"
        ))
        df["Sentiment Score"] = df["Feedback"].apply(lambda x: TextBlob(str(x)).sentiment.polarity)

        # Save the updated file to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)

        output.seek(0)

        # Return the processed Excel file as a downloadable response
        headers = {
            "Content-Disposition": "attachment; filename=updated_feedback.xlsx",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
        return Response(content=output.getvalue(), headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return {"error": str(e)}
