from google import genai
import os
from dotenv import load_dotenv

load_dotenv()

class GeminiAgent:

    def __init__(self):
        self.client = genai.Client(
            api_key=os.getenv("GEMINI_API_KEY")
        )

    def generate_content(self, prompt: str):
        response = self.client.models.generate_content(
            model="gemini-2.5-flash",   # or gemini-.5-pro
            contents=prompt
        )

        return response.text