# Document generation agent (A2A Compatible)

An A2A-compatible agent that autonomously generates structured .docx based on natural language briefs.

## Features
-Parses natural language input to create structured documents
-Generates well-organized content with headings and subheadings
-Formats paragraphs, bullet lists

## Setup

1. Install dependencies:
```bash
pip install -e .
```

2. Set up environment variables:
```bash
export OPENROUTER_API_KEY="your-api-key"
export MONGO_URL="mongodb://localhost:27017"
```

3. Ensure MongoDB is running for chat history storage.

## Running

```bash
python -m src --host localhost --port 10008
```

Options:
- `--host`: Host to bind to (default: localhost)
- `--port`: Port to bind to (default: 10008)
- `--mongo-url`: MongoDB connection URL (default: mongodb://localhost:27017)
- `--db-name`: Database name (default: compliance-checker-a2a)

## Usage

The agent exposes the following tools:

### check_compliance
Analyze a document for policy compliance.

**Parameters:**
- `document_text` (str): The document text to analyze
- `query` (str, optional): Specific question about compliance

### analyze_policy
Answer questions about specific policies.

**Parameters:**
- `policy_question` (str): Question about policies or compliance requirements

## Example Queries

- "Check this document for policy compliance"
- "Does this email violate any policies?"
- "What are the encryption requirements for file transfers?"
- "Analyze this expense report for compliance issues"
