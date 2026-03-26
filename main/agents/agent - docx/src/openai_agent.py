from docx_toolset import DocxToolset


def create_agent(host: str, port: int):
    toolset = DocxToolset(host, port)
    return {
        "tools": toolset.get_tools(),
        "system_prompt": """You are a Document Architect Agent.
        Your goal is to convert natural language requests into structured Word (.docx) files.
        
        1. Parse the user request to identify a Title and logical Sections (Headings and Content).
        2. Call the 'generate_docx' tool with a clear filename, the title, and the list of sections.
        3. Once the tool returns a URL, provide it to the user so they can download their file."""
    }
