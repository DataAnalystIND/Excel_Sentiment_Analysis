from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
from textblob import TextBlob
from fastapi.responses import StreamingResponse

app = FastAPI()

# ✅ CORS Middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    try:
        # Read the uploaded file
        contents = await file.read()
        df = pd.read_excel(BytesIO(contents), engine='openpyxl')

        # ✅ Ensure Column A Exists
        if df.shape[1] < 1:
            return {"error": "Column A (Feedback) is missing"}

        # ✅ Rename Column A for clarity
        df.rename(columns={df.columns[0]: "Feedback"}, inplace=True)

        # ✅ Perform Sentiment Analysis
        df["Sentiment"] = df["Feedback"].apply(lambda x: TextBlob(str(x)).sentiment.polarity)
        df["Sentiment Label"] = df["Sentiment"].apply(lambda x: "Positive" if x > 0 else "Negative" if x < 0 else "Neutral")

        # ✅ Save the processed file correctly
        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        # ✅ Return file as a download response
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=processed_feedback.xlsx"}
        )

    except Exception as e:
        return {"error": str(e)}
