# 📄 Document Generation Agent (DOCX) — A2A Compatible

An **A2A-compatible AI agent** that converts structured input (JSON or natural language) into fully formatted **.docx documents** with headings, paragraphs, lists, and charts.

---

## 🚀 Features

* ✅ Accepts **JSON or natural language input**
* ✅ Generates **well-structured DOCX documents**
* ✅ Supports:

  * Headings & subheadings
  * Paragraphs
  * Bullet & numbered lists
  * **Bold** and *italic* formatting
* ✅ Docker support for easy deployment

---

## 📁 Project Structure

```
agent-docx/
│
├── src/
│   ├── __main__.py        # Entry point
│   ├── agent.py           # Agent controller
│   ├── executor.py        # Execution logic
│   ├── parser.py          # JSON input parser
│   ├── tools.py           # DOCX generation logic
│   └── sample.json        # Sample input file
│
├── requirements.txt
└── Dockerfile
```

---

## ⚙️ Installation

### 🔹 1. Clone the repository

```
git clone <your-repo-url>
cd agent-docx
```

---

### 🔹 2. Install dependencies

```
pip install -r requirements.txt
```

---

## 🧾 Input Format

### ✅ Simple JSON Input

```
{
  "title": "AI Report",
  "content": "Artificial Intelligence is transforming the world.\n\n**Machine Learning** is powerful.\n*Deep Learning* is a subset of AI."
}
```

---
## ▶️ Usage

### 🔹 Run Locally

```
cd src
python __main__.py
```

---

### 🔹 Run with Custom Input

```
python __main__.py sample.json
```

---

## 📄 Output

* Generates a `.docx` file:

```
output.docx
```

### Includes:

* Document title
* Structured sections
* Lists and paragraphs
* Rich text formatting

---
## 🔌 A2A Compatibility

This agent is compatible with **A2A (Agent-to-Agent)** architecture:

* Modular tool-based design
* Easily extendable for:

  * APIs
  * Multi-agent workflows
  * LLM integrations (OpenAI / Gemini)

---

## ⭐ Conclusion

This agent enables **automated document generation pipelines**, making it ideal for:

* Reports
* Assignments
* Business documents
* Data summaries

---

🚀 *Turn structured input into professional documents instantly!*
