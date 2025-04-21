from fastapi import FastAPI
from app.routers import tasks

app = FastAPI()

app.include_router(tasks.router)

@app.get("/")
def read_root():
    return {"message": "Welcome to the Automation App"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)