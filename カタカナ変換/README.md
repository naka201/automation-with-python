### カタカナ変換
業者様から送られてくる出品表の中の色や商品名の外国語を読み(カタカナ)に変換するプログラムです。

### メリット

1. **作業の効率化**  
   これまでは、外国語の読みを慣れている社員が出品表を印刷し、1つ1つ手作業で訂正していましたが、このプログラムを導入することで、**手作業を自動化し、数時間かかっていた作業を短時間で完了**できるようになりました。

2. **作業分担の柔軟化**  
   手作業では他の社員やアルバイトに教える時間が取れなかったため、この作業になれている社員が必ずその作業を担当する必要がありました。しかし、プログラムを使うことで、**実行方法さえ理解すれば、誰でも簡単に作業を引き継ぐことができ、作業負担の分散が可能**になりました。

3. **ヒューマンエラーの削減**  
   外国語の読みは、知らない単語や覚えていない単語が出てくるたびに調べて訂正していましたが、プログラムを使用することで、**誤りを減らし、正確なカタカナ変換を実現**できるようになりました。

4. **情報の共有と更新が容易に**  
   辞書を一度作成すれば、以後はその情報を再利用できます。辞書にない単語が出現した場合、プログラム実行後にその単語を辞書に追加するだけで、次回から自動的に変換されます。これにより、**読みの共有や更新が簡単に行えるようになりました**。


#### 必要ファイル
- 辞書用エクセルファイル
- 変換する出品表
- 出力用エクセルファイル

#### 注意事項  
- プログラムは、指定された出品表をカタカナに変換し、結果を出力用Excelファイルに保存します。  
- 辞書にない単語が出品表に含まれている場合、その単語は変換されません。  
- 辞書にない単語は、辞書用Excelファイルに追加しておく必要があります。新しい単語があれば適宜追加してください。  
- プログラムは自動化フォルダ内で以下のコマンドを実行することで動作します：  
  ```bash
  python3 カタカナ変換/main.py
