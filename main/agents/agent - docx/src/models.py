from pydantic import BaseModel

class DocxRequest(BaseModel):
    prompt: str

class DocxResponse(BaseModel):
    message: str
