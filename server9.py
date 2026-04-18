from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from openai import OpenAI
from dotenv import load_dotenv
import os
from typing import Tuple

# Load env
load_dotenv()

GROQ_API_KEY = os.getenv("GROQ_API_KEY")
HF_TOKEN = os.getenv("HF_TOKEN")

app = FastAPI(title="Simple LLM Switch Chatbot")


# ----------------------------
# Request schema
# ----------------------------
class ChatRequest(BaseModel):
    message: str
    model_type: str  # "censored" or "uncensored"


# ----------------------------
# Model switch
# ----------------------------
def get_model_client(model_type: str) -> Tuple[OpenAI, str]:
    if model_type == "censored":
        return (
            OpenAI(
                base_url="https://api.groq.com/openai/v1",
                api_key=GROQ_API_KEY,
            ),
            "llama-3.3-70b-versatile",
        )

    return (
        OpenAI(
            base_url="https://router.huggingface.co/v1",
            api_key=HF_TOKEN,
        ),
        "dphn/Dolphin-Mistral-24B-Venice-Edition:featherless-ai",
    )


# ----------------------------
# Chat endpoint
# ----------------------------
@app.post("/chat")
def chat(req: ChatRequest):
    try:
        client, model_name = get_model_client(req.model_type)

        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": "You are a helpful chatbot."},
                {"role": "user", "content": req.message},
            ],
            temperature=0.7,
            max_tokens=500,
        )

        return {
            "model_used": model_name,
            "reply": response.choices[0].message.content
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ----------------------------
# Health check
# ----------------------------
@app.get("/")
def root():
    return {"message": "Chatbot API is running"}
