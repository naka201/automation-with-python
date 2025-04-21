from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.routers import tasks
from app.api.v1.endpoints.settlement import router as settlement_router

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ルーターを登録
app.include_router(tasks.router, prefix="/tasks", tags=["tasks"])
app.include_router(settlement_router, prefix="/api/v1/settlement", tags=["settlement"])

@app.get("/")
def read_root():
    return {"message": "Welcome to the Unified Project!"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)