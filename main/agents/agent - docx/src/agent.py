from src.parser import parse_input
from src.tools import create_docx


class DocxAgent:
    def run(self, input_file="input.json"):
        title, sections = parse_input(input_file)

        output_file = create_docx(title, sections)

        return f"✅ Document created: {output_file}"