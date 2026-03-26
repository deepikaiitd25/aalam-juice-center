import os
import json
import logging
from docx import Document

logger = logging.getLogger(__name__)


class DocxToolset:
    def __init__(self, host: str, port: int):
        self.host = host
        self.port = port
        self.output_dir = "outputs"
        os.makedirs(self.output_dir, exist_ok=True)

    # THE FIX: Made async and returning a clean string for the A2A protocol
    async def generate_docx(self, filename: str, title: str, sections: list) -> str:
        """Generate a Word document with specific sections."""
        try:
            if isinstance(sections, str):
                logger.info("Sections received as string, parsing JSON...")
                sections = json.loads(sections)

            if not filename.endswith(".docx"):
                filename += ".docx"

            doc = Document()
            doc.add_heading(title, 0)

            for section in sections:
                # Fallback in case Gemini hallucinates the structure
                if isinstance(section, dict):
                    heading = section.get("heading", "Section")
                    content = section.get("content", "")
                else:
                    heading = "Section"
                    content = str(section)

                doc.add_heading(heading, level=1)
                doc.add_paragraph(content)

            filepath = os.path.join(self.output_dir, filename)
            doc.save(filepath)

            url = f"http://{self.host}:{self.port}/outputs/{filename}"
            # Return a user-friendly string that the A2A TextPart will display
            return f"✅ Successfully generated **{filename}**!\n\n📥 [Download your report here]({url})"

        except Exception as e:
            logger.error(f"❌ DOCX Generation failed: {e}")
            return f"❌ Failed to generate report: {str(e)}"

    def get_tools(self):
        return {"generate_docx": self}
