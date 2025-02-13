### 差額チェック
時計・宝石大会において、落札金額とMax入札金額の差額が、各大会で決まった条件を超えていないかチェックするプログラムです。

#### メリット

1. **作業時間の大幅短縮**  
   以前は、1000~1500点ほどの商品の差額を異常値でないか1行ずつ手動で確認していたため、チェック作業に**数時間**以上かかっていました。しかし、このプログラムを実行することで、異常な差額が含まれるセルに自動的に色が付けられ、色が付いたセルのみをチェックすればよくなり、作業時間が**数分**で完了するようになりました。

2. **エラー発見の視覚化**  
   異常な差額のセルが自動で色付けされるため、視覚的に簡単に問題箇所を特定でき、**エラーの見逃しを防止**できます。これにより、チェック作業の精度が向上しました。

#### 必要ファイル
- 入札が集計されたExcelファイル


#### 注意事項
- 条件を超えたセルが黄色に塗られて、新たなExcelファイルとして出力されます。
- すべての準備が整ったら、以下のコマンドでプログラムを実行します：  
  ```bash
  python3 差額チェック/main.py
