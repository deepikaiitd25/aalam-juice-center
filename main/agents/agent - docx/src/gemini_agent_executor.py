from gemini_agent import GeminiAgent
from tools import DocumentTools

class AgentExecutor:

    def __init__(self):
        self.agent = GeminiAgent()
        self.tools = DocumentTools()

    def execute(self, user_input):
        content = self.agent.generate_content(user_input)

        return self.tools.generate_docx(
            title="Generated Document",
            content=content
        )