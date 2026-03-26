import logging
import os
from typing import List, Dict, Tuple, Optional

from langchain_community.vectorstores import FAISS
from langchain_google_genai import GoogleGenerativeAIEmbeddings

logger = logging.getLogger(__name__)


class VectorStoreError(Exception):
    pass


class VectorStoreService:
    def __init__(self):
        self.embeddings = self._create_embeddings()
        self._store_cache: Optional[FAISS] = None
        self._cache_hash: Optional[str] = None

    def _create_embeddings(self) -> GoogleGenerativeAIEmbeddings:
        api_key = os.getenv("GEMINI_API_KEY") or os.getenv("OPENAI_API_KEY")

        if not api_key:
            raise VectorStoreError(
                "CRITICAL: No API Key found in Environment for Embeddings.")

        logger.info(
            "Initializing Native Gemini Embeddings (gemini-embedding-001)...")
        try:
            # Using the EXACT model name authorized for your API key!
            return GoogleGenerativeAIEmbeddings(
                model="models/gemini-embedding-001",
                google_api_key=api_key
            )
        except Exception as e:
            raise VectorStoreError(
                f"Failed to initialize Native Gemini Embeddings: {str(e)}")

    def create_vector_store(self, agent_cards: List[Dict[str, str]], use_cache: bool = True) -> FAISS:
        cards_hash = self._hash_agent_cards(agent_cards)
        if use_cache and self._is_cache_valid(cards_hash):
            return self._store_cache

        texts, metadatas = self._prepare_data(agent_cards)
        if not texts:
            raise VectorStoreError("No valid agent cards found")

        vectorstore = FAISS.from_texts(
            texts, embedding=self.embeddings, metadatas=metadatas)
        self._store_cache = vectorstore
        self._cache_hash = cards_hash
        return vectorstore

    def _prepare_data(self, agent_cards: List[Dict[str, str]]):
        texts, metadatas = [], []
        for card in agent_cards:
            desc = card.get("description", "")
            name = card.get("name", "")
            if desc and name:
                texts.append(desc)
                metadatas.append({"name": name})
        return texts, metadatas

    def _hash_agent_cards(self, agent_cards):
        import hashlib
        import json
        sorted_cards = sorted(agent_cards, key=lambda x: x.get("name", ""))
        return hashlib.md5(json.dumps(sorted_cards, sort_keys=True).encode()).hexdigest()

    def _is_cache_valid(self, cards_hash):
        return self._store_cache and self._cache_hash == cards_hash

    def similarity_search(self, vectorstore: FAISS, query: str, k: int = 5):
        results = vectorstore.similarity_search_with_score(query, k=k)
        return [{"name": doc.metadata["name"], "similarity_score": float(score)} for doc, score in results]

    def clear_cache(self):
        self._store_cache = self._cache_hash = None
