from fastapi import APIRouter
from app.services.task_service import get_tasks

router = APIRouter()

@router.get("/")
def get_tasks_endpoint():
    """利用可能な自動化業務の一覧を取得"""
    return get_tasks()

