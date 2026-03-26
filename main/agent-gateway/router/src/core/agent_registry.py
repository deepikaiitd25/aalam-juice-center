import logging
from typing import List, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


class AgentRegistryError(Exception):
    pass


class AgentRegistry:
    """Mock registry for local hackathon testing without a backend database."""

    def __init__(self):
        pass

    async def fetch_agent_cards(self, token: str) -> List[Dict[str, str]]:
        logger.info(
            "🛠️ [HACKATHON MODE] Bypassing backend, loading hardcoded agents...")

        return [
            {
                "id": "excel_agent",
                "name": "Excel Generation Agent",
                "description": "Autonomously generates structured .xlsx spreadsheets, tabular forms, and data grids.",
                "url": "http://localhost:10008/"
            },
            # UNCOMMENT THIS WHEN THE PPTX AGENT IS READY:
            # {
            #     "id": "pptx_agent",
            #     "name": "PowerPoint Generation Agent",
            #     "description": "Generates .pptx presentation slides and pitch decks.",
            #     "url": "http://localhost:10009/"
            # }
        ]

    def get_agent_url(self, agent_cards: List[Dict[str, str]], agent_name: str) -> Optional[str]:
        for card in agent_cards:
            if card.get("name") == agent_name:
                return card.get("url")
        return None

    def get_fallback_agent(self, agent_cards: List[Dict[str, str]]) -> Optional[Tuple[str, str]]:
        if agent_cards:
            return agent_cards[0].get("name"), agent_cards[0].get("url")
        return None
