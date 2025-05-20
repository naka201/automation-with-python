# F-project/app/main.py

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.routers import tasks # tasks.py が app/routers/ にあると仮定します
from app.api.v1.endpoints.settlement import router as settlement_router
from app.routers.brand_search import router as brand_search_router # ブランド検索ルーターをインポート

# FastAPIアプリケーションのインスタンスを作成
# title, description, version などを設定すると /docs で表示が見やすくなります
app = FastAPI(
    title="Unified Project API",
    description="This API provides functionalities for tasks, settlement processing, and brand search.",
    version="1.0.0"
)

# CORS (Cross-Origin Resource Sharing) ミドルウェアの設定
# 全てのオリジンを許可していますが、本番環境ではセキュリティのため具体的なオリジンを指定してください。
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 例: ["http://localhost:3000", "https://your-frontend-domain.com"]
    allow_credentials=True,
    allow_methods=["*"], # GET, POST, PUT, DELETE など許可するメソッド
    allow_headers=["*"], # 許可するHTTPヘッダー
)

# --- ルーターの登録 ---

# 既存のルーター (例: tasks)
# `tasks.router` が `app.routers.tasks` モジュールで定義されていると仮定
if 'tasks' in globals() and hasattr(tasks, 'router'): # tasksモジュールとrouterの存在確認
    app.include_router(tasks.router, prefix="/tasks", tags=["tasks"])
else:
    print("Warning: tasks router not found or not imported correctly.")


# 既存のルーター (例: settlement)
app.include_router(settlement_router, prefix="/api/v1/settlement", tags=["settlement"])


# ブランド検索ルーターの登録
# brand_search_router 内のエンドポイント (例: /process-files/) が
# /api/v1/brand-search/process-files/ のようにアクセスできるようにprefixを設定します。
app.include_router(
    brand_search_router,
    prefix="/api/v1/brand-search",  # 推奨されるprefix
    tags=["brand-search"]           # タグは既存のAPIと一貫性を持たせる
)

# ルートエンドポイント
@app.get("/")
async def read_root(): # FastAPIでは async def を使用することが一般的です
    return {"message": "Welcome to the Unified Project API! Navigate to /docs for the API documentation."}

# このifブロックは `python app/main.py` で直接実行した際にUvicornサーバーを起動します。
# 本番環境や開発時でも、通常はプロジェクトルートから `uvicorn app.main:app --reload` コマンドで起動します。
if __name__ == "__main__":
    import uvicorn
    # 開発時はリロードを有効にすると便利です。
    # uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
    # もし `app` ディレクトリ内で `python main.py` を実行する場合は以下
    uvicorn.run(app, host="0.0.0.0", port=8000)