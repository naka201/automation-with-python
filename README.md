# automation-with-python

## 概要
これは業務の自動化を目的としたプロジェクトです。
Excelデータの操作やMicrosoft Access, Outlookとの連携を通じて、データ処理やメール送信業務を効率化することを目指しています。

## 各業務説明
このプロジェクトで自動化した業務は以下の7つの業務です。

- **カタカナ変換**　：出品表の外国語をカタカナに変換するプログラム。
- **メール送信**　　：精算書を送るメールを自動で生成するプログラム。
- **入札貼付**　　　：大会毎に集まった出品表を1つのエクセルファイルに貼り付けるプログラム。
- **差額チェック**　：集計した入札のある金額の差額が大会ごとに決まった条件を超えていないかチェックするプログラム。
- **店舗出品リスト**：全体の出品表から必要な情報を抽出して箱番ごとの出品表を作成するプログラム。
- **検品チェック**　：検品と入力の管理を行うプログラム。
- **箱番作成**　　　：出品者に自動で箱番を割り当てるプログラム。


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
- pywin32 ：outlookとの連携
- pyyaml  ：YAMLファイルの読み込み
- pandas  ：データフレーム
- openpyxl：エクセルファイルの読み書き
- pyodbc  ：Microsoft Accessとの接続

