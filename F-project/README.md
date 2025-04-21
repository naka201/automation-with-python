# Unified Automation and Settlement Project

このプロジェクトは、FastAPIを使用して以下の2つの主要な機能を提供します：
1. 自動化業務の管理と実行
2. 購入者および出品者ごとの精算書の生成

## プロジェクト構成

```
unified-project/
    app/
        main.py                # アプリケーションのエントリポイント
        routers/
            tasks.py           # 自動化業務のルーター
        services/
            task_service.py    # 自動化業務のビジネスロジック
            settlement_service.py # 精算書生成のビジネスロジック
        schemas/
            __init__.py        # データスキーマの定義
        api/
            v1/
                endpoints/
                    settlement.py # 精算書生成APIのエンドポイント
        core/
            config.py          # アプリケーション設定
    requirements.txt           # プロジェクトの依存関係
    README.md                  # プロジェクトの説明
```

---

## 機能

### 1. 自動化業務の管理
- 利用可能な自動化業務の一覧を取得
- 特定の業務を実行

#### エンドポイント
- **GET /tasks**: 利用可能な自動化業務の一覧を取得します。

#### 使用例
```bash
curl http://localhost:8000/tasks
```

#### レスポンス例
```json
[
    {"id": 1, "name": "精算書作成", "description": "精算書を作成するタスク"},
    {"id": 2, "name": "データ集計", "description": "データを集計するタスク"},
    {"id": 3, "name": "レポート生成", "description": "レポートを生成するタスク"}
]
```

---

### 2. 精算書の生成
- 購入者および出品者ごとの精算書を生成
- 精算書をZIPファイルとしてダウンロード可能

#### エンドポイント
- **POST /api/v1/settlement/generate**: 精算書を生成します。

#### 使用例
```bash
curl -X POST "http://localhost:8000/api/v1/settlement/generate" \
-F "input_file=@input.xlsx" \
-F "template_buyer=@buyer_template.xlsx" \
-F "template_seller=@seller_template.xlsx" \
-F "日付=2023-10-01" \
-F "担当=山田太郎"
```

#### レスポンス
生成された精算書を含むZIPファイルが返されます。

---

## セットアップ

### 1. リポジトリをクローン
```bash
git clone <repository-url>
cd unified-project
```

### 2. 必要な依存関係をインストール
```bash
pip install -r requirements.txt
```

### 3. アプリケーションを起動
```bash
uvicorn app.main:app --reload
```

---

## 使用方法

1. ブラウザまたはAPIクライアントを使用して、以下のURLにアクセスします：
   - 自動化業務の一覧: `http://localhost:8000/tasks`
   - 精算書生成: `http://localhost:8000/api/v1/settlement/generate`

2. 必要に応じて、エンドポイントにリクエストを送信して機能を利用します。

---

## ライセンス

このプロジェクトはMITライセンスの下で提供されています。