# ダミーデータ作成ツール

高校新入生データベースのテスト用ダミーデータを生成するアプリケーションです。

## 機能

- 学生番号（2から始まる7桁）
- 名前（日本人名）
- 性別（男、女）
- 生年月日（高校新入生の学年に合わせた生年月日）
- 東京都の住所
- 保護者氏名（学生と同じ苗字）
- 保護者電話番号
- メールアドレス（2から始まる10桁@ドメイン）

## インストール

### 実行ファイルを使用する場合

[Releases](https://github.com/seyaytua/dummyData/releases)から、お使いのOSに対応したファイルをダウンロードしてください。

- Windows: `DummyDataGenerator.exe`
- macOS: `DummyDataGenerator`

### ソースコードから実行する場合

```bash
# リポジトリをクローン
git clone https://github.com/seyaytua/dummyData.git
cd dummyData

# 仮想環境を作成
python -m venv venv

# 仮想環境を有効化
# Mac/Linux:
source venv/bin/activate

# 依存関係をインストール
pip install -r requirements.txt

# アプリを実行
python main.py
使い方
「参照」ボタンをクリックして、データを保存するExcelファイルを選択または作成
生成したい項目にチェックを入れる
列名と件数を設定
「データ生成」ボタンをクリック
ビルド
Copy# macOS
pyinstaller --onefile --windowed --name "DummyDataGenerator" main.py
ライセンス
MIT License
