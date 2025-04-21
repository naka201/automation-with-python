from fastapi import FastAPI

app = FastAPI()

from .routers import tasks

app.include_router(tasks.router)