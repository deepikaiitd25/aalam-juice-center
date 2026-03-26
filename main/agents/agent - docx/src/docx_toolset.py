import os
import json  # Added for parsing
import logging
from docx import Document
from pydantic import BaseModel

logger = logging.getLogger(__name__)


class DocxResponse(BaseModel):
    status: str
    file_url: str | None = None
    error: str | None = None


class DocxToolset:
    def __init__(self, host: str, port: int):
        self.host = host
        self.port = port
        self.output_dir = "outputs"
        os.makedirs(self.output_dir, exist_ok=True)

    def generate_docx(self, filename: str, title: str, sections: list) -> DocxResponse:
        try:
            # FIX: If sections arrived as a string, parse it into a list
            if isinstance(sections, str):
                logger.info("Sections received as string, parsing JSON...")
                sections = json.loads(sections)

            if not filename.endswith(".docx"):
                filename += ".docx"

            doc = Document()
            doc.add_heading(title, 0)

            for section in sections:
                # Double-check that section is actually a dict
                heading = section.get("heading", "Section")
                content = section.get("content", "")

                doc.add_heading(heading, level=1)
                doc.add_paragraph(content)

            filepath = os.path.join(self.output_dir, filename)
            doc.save(filepath)

            url = f"http://{self.host}:{self.port}/outputs/{filename}"
            return DocxResponse(status="success", file_url=url)
        except Exception as e:
            logger.error(f"❌ DOCX Generation failed: {e}")
            return DocxResponse(status="error", error=str(e))

    def get_tools(self):
        return {"generate_docx": self}
