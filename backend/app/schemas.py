from pydantic import BaseModel,Field
class LoginRequest(BaseModel):
    username:str=Field(min_length=1,max_length=100)
    password:str=Field(min_length=1,max_length=200)
class LoginResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in_minutes: int
    username: str
    role: str

class HealthResponse(BaseModel):
    status:str
    service:str
