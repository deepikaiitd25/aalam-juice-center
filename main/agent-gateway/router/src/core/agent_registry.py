import logging
from typing import List, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


class AgentRegistryError(Exception):
    """Custom exception for agent registry errors."""
    pass


class AgentRegistry:
    """Mock registry for local hackathon testing without a backend database."""

    def __init__(self):
        pass

    async def fetch_agent_cards(self, token: str) -> List[Dict[str, str]]:
        """
        Hardcoded registry for local services.
        Bypasses backend database for the hackathon environment.
        """
        logger.info(
            "🛠️ [HACKATHON MODE] Bypassing backend, loading hardcoded agents...")

        return [
            {
                "id": "docx_agent",
                "name": "Document Generation Agent",
                "description": "Autonomously generates structured .docx files, formal reports, and professional letters.",
                "url": "http://localhost:10007/"
            },
            {
                "id": "excel_agent",
                "name": "Excel Generation Agent",
                "description": "Autonomously generates structured .xlsx spreadsheets, tabular forms, and data grids.",
                "url": "http://localhost:10008/"
            },
            {
                "id": "pptx_agent",
                "name": "PowerPoint Generation Agent",
                "description": "Autonomously generates structured .pptx slide decks from a natural language brief.",
                "url": "http://localhost:10009/"
            }
        ]

    def get_agent_url(self, agent_cards: List[Dict[str, str]], agent_name: str) -> Optional[str]:
        """Find the URL for a specific agent by its name."""
        for card in agent_cards:
            if card.get("name") == agent_name:
                return card.get("url")
        return None

    def get_fallback_agent(self, agent_cards: List[Dict[str, str]]) -> Optional[Tuple[str, str]]:
        """Return the default agent if no specific routing is determined."""
        if agent_cards:
            return agent_cards[0].get("name"), agent_cards[0].get("url")
        return None
