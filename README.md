# automation-with-python

## 概要
これは業務の自動化を目的としたプロジェクトです。
Excelデータの操作やMicrosoft Access, Outlookとの連携を通じて、データ処理やメール送信業務を効率化することを目指しています。

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

