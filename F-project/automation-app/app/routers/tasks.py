from fastapi import APIRouter
from app.services.task_service import get_available_tasks

router = APIRouter()

@router.get("/tasks")
def get_tasks():
    """利用可能な自動化業務の一覧を取得"""
    return get_available_tasks()