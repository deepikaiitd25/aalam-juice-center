from agent import DocxAgent

if __name__ == "__main__":
    agent = DocxAgent()

    print("🚀 DOCX AI Agent Running...\n")

    while True:
        user_input = input("Enter topic: ")
        result = agent.run(user_input)
        print(result)