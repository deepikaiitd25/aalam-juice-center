import argparse
from agent import DocxAgent

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", default=5000)
    args = parser.parse_args()

    agent = DocxAgent()

    print("DOCX Agent running...")

    while True:
        user_input = input("Enter topic: ")
        result = agent.run(user_input)
        print(result)

if __name__ == "__main__":
    main()
