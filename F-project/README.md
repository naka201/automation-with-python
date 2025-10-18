# Automation with FastAPI

このプロジェクトは、FastAPIを使用して以下の2つの主要な機能を提供します：
1. 購入者および出品者ごとの精算書の生成
2. ブランド検索とハイライト機能

## プロジェクト構成

```
F-project/
    app/
        main.py                # アプリケーションのエントリポイント
        routers/
            tasks.py           # 自動化業務のルーター
            brand_search.py    # ブランド検索のルーター
        services/
            task_service.py    # 自動化業務のビジネスロジック
            settlement_service.py # 精算書生成のビジネスロジック
            brand_search_service.py # ブランド検索のビジネスロジック
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

### 1. ブランド検索とハイライト機能
- Excelファイルからブランドリストを読み込み
- PDFファイル内でブランドを検索
- 検索結果をExcelファイルにハイライト表示

#### エンドポイント
- **POST /process-files/**: ブランド検索とハイライト処理を実行します。

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
ハイライト処理されたExcelファイルが返されます。

---

## セットアップ

### 1. リポジトリをクローン
```bash
git clone <repository-url>
cd F-project
```

### 2. 必要な依存関係をインストール
```bash
pip install -r requirements.txt
```

### 3. アプリケーションを起動
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
uvicorn app.main:app --host <IPアドレス> --port 8000 --reload
```

---

## 使用方法

1. ブラウザまたはAPIクライアントを使用して、以下のURLにアクセスします：
   - APIドキュメント: `http://localhost:8000/docs`
   - APIドキュメント: `http://<IPアドレス>:8000/docs`

2. 実行したい業務のエンドポイントにリクエストを送信して機能を利用します。

