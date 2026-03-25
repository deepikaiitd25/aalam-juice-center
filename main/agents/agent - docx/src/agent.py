from gemini_agent_executor import AgentExecutor

class DocxAgent:

    def __init__(self):
        self.executor = AgentExecutor()

    def run(self, user_input):
        return self.executor.execute(user_input)