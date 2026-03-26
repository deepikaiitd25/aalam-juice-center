import logging
import os
import click
import uvicorn

from a2a.server.apps import A2AStarletteApplication
from a2a.server.request_handlers import DefaultRequestHandler
from a2a.server.tasks import InMemoryTaskStore
from a2a.types import AgentCapabilities, AgentCard, AgentSkill
from dotenv import load_dotenv
from starlette.applications import Starlette
from starlette.middleware.cors import CORSMiddleware
from starlette.routing import Mount
from starlette.staticfiles import StaticFiles

from openai_agent import create_agent
from openai_agent_executor import OpenAIAgentExecutor

load_dotenv()
logging.basicConfig(level=logging.INFO)


@click.command()
@click.option("--host", "host", default="localhost")
@click.option("--port", "port", default=10008)
@click.option("--mongo-url", "mongo_url", default="mongodb://localhost:27017")
@click.option("--db-name", "db_name", default="excel-agent-a2a")
def main(host: str, port: int, mongo_url: str, db_name: str):

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY environment variable must be set in .env")

    # Define the skill for Nasiko
    skill = AgentSkill(
        id="generate_xlsx",
        name="Generate Excel Spreadsheet",
        description="Generates a structured .xlsx spreadsheet based on a brief and data.",
        tags=["excel", "spreadsheet", "data"],
        examples=[
            "Build a sales performance tracker for Q1 with columns for rep name, deals closed, revenue, and quota attainment.",
        ],
    )

    agent_card = AgentCard(
        name="Excel Generation Agent",
        description="Autonomously generates structured .xlsx spreadsheets.",
        url=f"http://{host}:{port}/",
        version="1.0.0",
        default_input_modes=["text"],
        default_output_modes=["text"],
        capabilities=AgentCapabilities(streaming=True),
        skills=[skill],
    )

    # Pass host/port so the toolset can build accurate download URLs
    agent_data = create_agent(host=host, port=port)

    agent_executor = OpenAIAgentExecutor(
        card=agent_card,
        tools=agent_data["tools"],
        api_key=api_key,
        system_prompt=agent_data["system_prompt"],
        base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
        model="gemini-2.5-flash",
    )

    request_handler = DefaultRequestHandler(
        agent_executor=agent_executor, task_store=InMemoryTaskStore()
    )

    a2a_app = A2AStarletteApplication(
        agent_card=agent_card, http_handler=request_handler
    )

    # Get base routes and mount the static outputs directory
    os.makedirs("outputs", exist_ok=True)
    routes = a2a_app.routes()
    routes.append(
        Mount("/outputs", app=StaticFiles(directory="outputs"), name="outputs")
    )

    app = Starlette(routes=routes)

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    main()
