### メール送信  
このプログラムは、オークション大会後の精算書を送付するためのメールを自動で作成します。  
送信するPDFファイル名に業者コードを指定することで、Excelに保存された連絡先一覧から対応するメールアドレスを取得し、精算書を添付したメールの下書きを自動的に作成します。

### メリット

1. **セキュリティ面の強化**  
   精算書は社外秘の重要な情報であるため、誤って他社に送信するリスクを避けることが必須です。このプログラムは、精算書を送信する際のヒューマンエラーを防止し、誤送信による情報漏洩を未然に防ぐため、**セキュリティ面**でも大きなメリットを提供します。

2. **業務の効率化**  
   精算書の送信メール作成にかかる時間を大幅に削減でき、業務の**効率化**にも貢献します。業者コードを指定するだけで、対応するメールアドレスが自動的に取得され、添付ファイルとともに正確なメールが作成されるため、**送信ミス**のリスクを低減します。


#### 必要なファイル
- 送信するPDFファイル  
- メールアドレスをまとめたExcelファイル  
- メール本文のテンプレート（YAML形式）

#### 注意事項
- 送信するPDFファイルは、各大会ごとのフォルダに格納してください。  
- 大会当日に精算書を送信できない場合は、YAML形式のテンプレートファイルで大会の日付を指定してください。  
- すべての準備が整ったら、以下のコマンドでプログラムを実行します：  
  ```bash
  python3 メール送信/main.py
