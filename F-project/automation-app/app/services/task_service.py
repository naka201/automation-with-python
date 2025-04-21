from typing import List

TASKS = [
    {"id": 1, "name": "精算書作成", "description": "精算書を作成するタスク"},
    {"id": 2, "name": "データ集計", "description": "データを集計するタスク"},
    {"id": 3, "name": "レポート生成", "description": "レポートを生成するタスク"},
]

def get_tasks() -> List[dict]:
    return TASKS

def execute_task(task_id: int) -> str:
    task = next((task for task in TASKS if task["id"] == task_id), None)
    if task:
        return f"タスク '{task['name']}' を実行しました。"
    else:
        return "指定されたタスクは存在しません。"