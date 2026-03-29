# DGS モジュール棚卸ツール

D-Link DGS シリーズスイッチに SSH 接続し、全ポートの SFP モジュール情報を自動収集して Excel に出力するツールです。

## 背景・用途

手作業での機器調査は台数が増えると時間がかかりミスが生じやすいため本ツールを作成。実務のネットワーク更改案件にて使用し、モジュール在庫把握・発注数量確認・設計資料作成に活用。

## 機能

- 複数スイッチへの一括 SSH 接続
- 全ポートのリンク状態・モジュール型番・規格・速度・Duplex・接続先 Description を取得
- SFP モジュールの純正／非純正判定
- 接続方式（銅線／マルチモード光／シングルモード光）の自動判定
- Excel 出力（総合集計シート＋ホスト別詳細シート）
- 更改後の規格を入力するプルダウン＋条件付き書式（色付け）
- 規格別カウント集計（COUNTIF）の自動挿入

## 対応モジュール

| 型番 | 規格 |
|---|---|
| DEM-431XT | 10GBASE-SR |
| DEM-432XT | 10GBASE-LR |
| DEM-410T | 10GBASE-T |
| DEM-311GT | 1000BASE-SX |
| DEM-310GT | 1000BASE-LX |
| DEM-330T | 1000BASE-BX-D |
| DEM-330R | 1000BASE-BX-U |
| DGS-712 | 1000BASE-T |
| DEM-211 | 100BASE-FX |

## 必要なライブラリ

```
pip install netmiko pandas openpyxl
```

## 設定

`get_dlink_modules.py` の以下の箇所を環境に合わせて書き換えてください。

```python
HOSTS = [
    {"hostname": "該当ホスト名_1", "ip": "該当IPアドレス_1"},
    ...
]
```

```python
def get_password(hostname):
    if hostname.startswith("ES"):
        return "該当パスワード(ESシリーズ)"
    elif hostname.startswith("BS"):
        return "該当パスワード(BSシリーズ)"
```

`collect_switch_data` 関数内の `username` も変更してください。

```python
device = {
    'username': '該当ユーザー名',
    ...
}
```

## 使い方

```bash
python get_dlink_modules.py
```

実行後、`DGS_Module_Inventory_Complete_YYYYMMDD_HHMMSS.xlsx` が同フォルダに生成されます。

### テスト実行（1台のみ）

`test_single_host.py` を使うと1台だけを対象にして動作確認できます。

```bash
python test_single_host.py
```

## 出力 Excel の構成

### 00_総合集計シート

全ホストのモジュール型番別の搭載数を一覧表示します。

| ホスト名 | DEM-431XT | DEM-311GT | ... | 空きスロット |
|---|---|---|---|---|
| 該当ホスト名_1 | 2 | 10 | ... | 5 |

### ホスト別詳細シート

各ポートの詳細情報を表示します。

| ポート番号 | リンク状態 | 接続方式 | モジュール型番 | 規格 | 速度 | Duplex | Type | 接続先 | 更改後の規格 |
|---|---|---|---|---|---|---|---|---|---|

- **J列（更改後の規格）** にプルダウンで規格を選択できます
- 選択した規格に応じて自動で色が変わります

### 色の凡例

| 色 | 規格 |
|---|---|
| 黄色 | 10GBASE-SR |
| 水色 | 10GBASE-LR |
| 薄い黄色 | 1000BASE-SX |
| 薄い紫 | 100BASE-FX |
| 薄い橙 | 1000BASE-T |

## 注意事項

- SSH 接続に netmiko を使用しています。device_type は `cisco_ios` を指定しています（DGS シリーズで動作確認済み）
- SFP モジュールが `reading...` 状態の場合、1秒待機して再取得します
- 接続ログは `{ホスト名}_teraterm_log.txt` として保存されます
