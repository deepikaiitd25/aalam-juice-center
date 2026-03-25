from openai_agent_executor import OpenAIAgentExecutor

class DocxAgent:

    def __init__(self):
        self.executor = OpenAIAgentExecutor()

    def run(self, user_input):
        return self.executor.execute(user_input)
