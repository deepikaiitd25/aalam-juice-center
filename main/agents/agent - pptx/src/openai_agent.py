import os
from openai import AsyncOpenAI
from .compliance_toolset import ComplianceToolset

class OpenAIAgent:
    def __init__(self):
        self.client = AsyncOpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        self.toolset = ComplianceToolset()
        self.system_prompt = (
            "You are a Presentation Assistant. Your goal is to help users create PowerPoints. "
            "When a user asks for a presentation, use the 'create_presentation' tool. "
            "Structure the content logically into slides before calling the tool."
        )

    async def process_message(self, text: str):
        # Define the tool for OpenAI
        tools = [
            {
                "type": "function",
                "function": {
                    "name": "create_presentation",
                    "description": "Generates a PowerPoint file.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "presentation_title": {"type": "string"},
                            "slides_json": {"type": "string", "description": "JSON array of objects with 'title' and 'content'"}
                        },
                        "required": ["presentation_title", "slides_json"]
                    }
                }
            }
        ]

        response = await self.client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": text}
            ],
            tools=tools
        )
        
        return response.choices[0].message
