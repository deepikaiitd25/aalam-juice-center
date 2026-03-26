# PowerPoint Generation Agent (A2A Compatible)

An A2A-compatible agent that autonomously generates structured `.pptx` slide decks based on natural language briefs.

## Features
- Parses natural language briefs into structured slide content
- Supports 4 slide types: title, content (bullets), two-column comparison, closing
- 5 color themes: blue, green, dark, red, purple
- Serves generated files over HTTP so users can download directly
- Uses Gemini 2.5 Flash via OpenAI-compatible API

## File Structure

```
agent-pptx/
├── src/
│   ├── __main__.py              # Entry point, A2A server setup
│   ├── openai_agent.py          # Agent factory + system prompt
│   ├── openai_agent_executor.py # A2A executor loop
│   ├── pptx_toolset.py          # generate_pptx tool + slide builders
│   └── models.py                # A2A Pydantic models
├── outputs/                     # Generated .pptx files served here
├── Dockerfile
├── docker-compose.yml
├── pyproject.toml
├── AgentCard.json
└── .gitignore
```

## Setup & Run

### 1. Set your API key

Create a `.env` file in the project root:
```
GEMINI_API_KEY=your_gemini_api_key_here
```

### 2. Build and run with Docker

```bash
docker build -t a2a-pptx-agent .
docker run -p 5000:5000 -e GEMINI_API_KEY=your_key_here a2a-pptx-agent
```

### 3. Test with curl

```bash
curl -X POST http://localhost:5000/ \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "test-001",
    "method": "message/send",
    "params": {
      "session_id": "test-001",
      "message": {
        "role": "user",
        "parts": [{"kind": "text", "text": "Create a 8-slide pitch deck for a fintech startup"}],
        "messageId": "msg-001"
      }
    }
  }'
```

The agent will respond with a download URL like:
`http://localhost:5000/outputs/fintech_startup_pitch.pptx`

## Slide Types

| type | Required fields |
|------|----------------|
| `title` | title, subtitle |
| `content` | title, bullets (list), notes (optional) |
| `two_column` | title, left_title, left_bullets, right_title, right_bullets |
| `closing` | title, subtitle |

## Themes

`blue` (default) · `green` · `dark` · `red` · `purple`

## Deploy to Nasiko

```bash
zip -r agent-pptx.zip agent-pptx/ -x "*.pyc" "*/__pycache__/*" "*/.git/*" "*/.env" "*/outputs/*"
```
Then upload the ZIP via the Nasiko dashboard → Add Agent → Upload ZIP.
