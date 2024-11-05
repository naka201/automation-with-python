# automation-with-python

## 概要
これは業務の自動化を目的としたプロジェクトです。
Excelデータの操作やMicrosoft Access, Outlookとの連携を通じて、データ処理やメール送信業務を効率化することを目指しています。

## 各業務説明

### カタカナ変換
業者様から送られてくる出品表の中の色や商品名の外国語を読み(カタカナ)に変換するプログラムです。

#### 使用方法
カタカナ変換フォルダ内に、外国語と読みを対応させた辞書用のエクセルファイルと変換するエクセルファイル、出力先のエクセルファイル用意し、プログラムを実行する。  
プログラムは変換したエクセルファイルと、辞書にない単語を返すので、辞書にない単語が出品表に出てきた場合は、辞書用のエクセルファイルに適宜追加してください。


## インストール方法

1. このリポジトリをクローンします。

    ```bash
    git clone https://github.com/naka201/automation-with-python.git
    ```

2. 必要なパッケージをインストールします。

    プロジェクトには以下の依存関係があります。`requirements.txt` を使ってインストールできます。

    ```bash
    cd automation-with-python
    pip install -r requirements.txt
    ```

## 使用方法

スクリプトを実行するには、以下のコマンドを使います。
フォルダ名には、実行したい業務名を入力します。

```bash
python <フォルダ名>/main.py
```

## Requirements
- pywin32
- pyyaml
- pandas
- openpyxl
- pyodbc

