FROM python:3.11-slim

WORKDIR /app

# THE FIX: Copy everything into /app, preserving the 'src' folder structure
COPY . /app/

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

# THE FIX: Point Python to the file inside the preserved 'src' folder
CMD ["python", "src/__main__.py", "--host", "0.0.0.0", "--port", "5000"]