import os
import click
import uvicorn
from dotenv import load_dotenv
from starlette.applications import Starlette
from starlette.middleware.cors import CORSMiddleware
from starlette.routing import Mount
from starlette.staticfiles import StaticFiles

from a2a.server.apps import A2AStarletteApplication
from a2a.server.request_handlers import DefaultRequestHandler
from a2a.server.tasks import InMemoryTaskStore
from a2a.types import AgentCard, AgentCapabilities, AgentSkill

from openai_agent import create_agent
from openai_agent_executor import OpenAIAgentExecutor

# Load environment variables from .env file
load_dotenv()


@click.command()
@click.option("--host", default="0.0.0.0")
@click.option("--port", default=10007)
def main(host, port):
    # Retrieve the API key from environment variables
    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        print("🚨 Error: GEMINI_API_KEY not found in environment variables.")
        return

    # Define the skill with all required Pydantic fields
    skill = AgentSkill(
        id="docx_generation",
        name="DOCX Document Generation",
        description="Generates Microsoft Word (.docx) documents with headings, paragraphs, and lists.",
        tags=["ai", "docx", "word", "report", "automation"],
        examples=[
            "Create a report on Artificial Intelligence",
            "Generate a project document on climate change",
            "Make a structured document for business proposal"
        ]
    )

    # Create the Agent Card with the missing required fields
    card = AgentCard(
        name="docx-generator-agent",  # Matches your AgentCard.json
        description="An AI agent that converts natural language input into structured .docx documents.",
        version="1.0.0",  # Added: Required by Pydantic
        defaultInputModes=["text"],  # Added: Required by Pydantic
        defaultOutputModes=["text"],  # Added: Required by Pydantic
        url=f"http://{host}:{port}/",
        capabilities=AgentCapabilities(streaming=True),
        skills=[skill]
    )

    # Initialize the agent logic and tools
    agent_data = create_agent(host, port)

    # Initialize the executor
    executor = OpenAIAgentExecutor(
        card=card,
        tools=agent_data["tools"],
        api_key=api_key,
        system_prompt=agent_data["system_prompt"],
        base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
        model="gemini-2.5-flash"
    )

    # Setup the A2A request handler and application
    handler = DefaultRequestHandler(executor, InMemoryTaskStore())
    a2a_app = A2AStarletteApplication(card, handler)

    # Ensure the output directory exists
    os.makedirs("outputs", exist_ok=True)

    # Setup routes and mount the static 'outputs' folder
    routes = a2a_app.routes()
    routes.append(
        Mount("/outputs", app=StaticFiles(directory="outputs"), name="outputs"))

    # Create the Starlette app
    app = Starlette(routes=routes)
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["*"],
        allow_headers=["*"]
    )

    print(f"🚀 DOCX Agent starting on http://{host}:{port}")
    uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    main()
