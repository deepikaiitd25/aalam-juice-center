from openai import OpenAI

class OpenAIAgent:

    def __init__(self):
        self.client = OpenAI()

    def generate_content(self, prompt):
        response = self.client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Generate structured document content with headings and paragraphs."},
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content
