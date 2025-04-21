from pydantic import BaseModel

class TaskSchema(BaseModel):
    id: int
    name: str
    description: str
    status: str

class CreateTaskSchema(BaseModel):
    name: str
    description: str

class UpdateTaskSchema(BaseModel):
    name: str = None
    description: str = None
    status: str = None