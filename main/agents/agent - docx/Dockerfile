FROM python:3.11-slim

WORKDIR /app

# Copy source code
COPY src/ /app

# Install dependencies
RUN pip install --no-cache-dir \
    "a2a-sdk[http-server]>=0.3.0" \
    click>=8.1.8 \
    openai>=1.57.0 \
    pydantic>=2.11.4 \
    python-dotenv>=1.1.0 \
    uvicorn>=0.34.2 \
    python-docx \
    matplotlib \
    openpyxl

# Environment settings
ENV PYTHONUNBUFFERED=1

# Run your agent
CMD ["python", "__main__.py", "--host", "0.0.0.0", "--port", "5000"]
