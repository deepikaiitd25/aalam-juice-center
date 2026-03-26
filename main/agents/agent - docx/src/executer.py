from src.agent import DocxAgent


def execute():
    agent = DocxAgent()
    result = agent.run()

    print(result)