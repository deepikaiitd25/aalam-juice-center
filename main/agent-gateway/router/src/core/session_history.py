import logging
import httpx
from typing import Any, List

logger = logging.getLogger(__name__)

# Add this missing class back so the Router doesn't crash on import!


class SessionHistoryError(Exception):
    pass


class SessionHistoryService:
    """Service to connect directly to the Chat History Microservice."""

    def __init__(self):
        # Pointing directly to your new microservice on port 8002
        self.history_url = "http://localhost:8002/chat-history/"

    async def fetch_session_history(self, token: str, session_id: str) -> List[Any]:
        logger.info(
            f"📚 Fetching actual session history from {self.history_url}{session_id}")
        try:
            # We use httpx to make an async call to your Chat History service
            async with httpx.AsyncClient() as client:
                response = await client.get(f"{self.history_url}{session_id}")

                if response.status_code == 200:
                    data = response.json()
                    messages = data.get("messages", [])
                    logger.info(
                        f"✅ Found {len(messages)} past messages for this session.")
                    return messages
                else:
                    logger.warning(
                        f"⚠️ No history found (Status: {response.status_code})")
                    return []
        except Exception as e:
            logger.error(f"❌ Failed to connect to Chat History Service: {e}")
            return []

    def reconstruct_conversation(self, messages: List[Any]) -> str:
        """Formats the past JSON messages into a readable text script for the AI."""
        if not messages:
            return ""

        conversation = []
        for msg in messages:
            role = msg.get("role", "user")
            content = msg.get("content", "")
            conversation.append(f"{role.upper()}: {content}")

        return "\n".join(conversation)
