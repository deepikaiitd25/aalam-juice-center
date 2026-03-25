from openai_agent import OpenAIAgent
from compliance_toolset import DocumentToolset

class OpenAIAgentExecutor:

    def __init__(self):
        self.agent = OpenAIAgent()
        self.toolset = DocumentToolset()

    def execute(self, user_input):
        content = self.agent.generate_content(user_input)
        title = "Generated Document"

        return self.toolset.generate_docx(title, content)
