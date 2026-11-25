# tax-system

古物台帳と輸出・免税販売の仕入販売相対表を、既存フォーマットのまま管理・検証・再出力するローカルシステムです。

実データ、テンプレートの複製、SQLiteデータベース、出力ファイルは `.tax-system/` または `outputs/` に保存され、Git管理から除外されます。原本は読み取り専用で参照し、直接更新しません。

## 運用しているPC

2026-11 に前のノートPC（Windows）からこのPC（Windows + WSL の Ubuntu）へ移管した。**今後はこのPCだけで使う**。前のノートPCでは開かない（データが2つに分かれるため）。

- 置き場所: `~/projects/tax-system`（WSL側）
- 実データ: `~/projects/tax-system/.tax-system/`（Git管理外。前のPCの `C:\Users\user\.tax-system` をコピーしたもの）
- テンプレートの保存パスは前のPCの `C:\...` のままDBに残っているが、このPCの `.tax-system/templates/` に自動で読み替える

## セットアップ（初回だけ。このPCでは済んでいる）

```bash
cd ~/projects/tax-system
python3 -m venv .venv
. .venv/bin/activate
pip install -e .
tax-system init
```

## Web画面（非エンジニア向け）

```bash
cd ~/projects/tax-system
. .venv/bin/activate
tax-system-web
```

起動後、Windows のブラウザで `http://127.0.0.1:5000` を開くと（WSL のサーバーにそのままつながる）、コマンド入力なしで以下が行えます。

- テンプレート登録・古物台帳/相対表/期末在庫表の取込
- 取込データの中身の閲覧（ページ送り対応）
- チェック結果の確認、確認用／正式出力
- **仕入との自動対応付け**: 相対表の各販売行に対応する古物台帳の購入記録を、商品名・数量・日付・取引相手から自動で探して割り当てる。1件の仕入は1件の販売にしか使わない。候補が複数・見つからない場合は画面から手動で選択
- **古物台帳の内訳復元**: 買取合計金額のみで商品明細が空欄の行について、相対表からわかる内訳を自動反映し、残額は期末在庫表の仕入れ原価から手動で埋めて内訳付きの古物台帳を出力

サーバーはこのPC内（127.0.0.1）だけで待ち受け、外部には公開されません。

## 基本操作（CLI）

`. .venv/bin/activate` してから実行する。Windows 側のファイルは `/mnt/c/Users/<ユーザー名>/Downloads/...` のように指定できる。

```bash
# テンプレートを版管理領域へ登録
tax-system register-template --type comparison --file "相対表.xlsx" --name optcg --version 2025-2026 --effective-from 2025-06-01

# 古物台帳CSVまたは相対表Excelを取り込む（返された取込IDを控える）
tax-system import --type ledger --file "approval_histories.csv"
tax-system import --type comparison --file "相対表.xlsx"

# 正式出力前チェック
tax-system validate 1

# エラー時に利用できる確認用出力
tax-system export 2 --template-id 2 --output "outputs/相対表_確認用.xlsx" --preview

# エラーが0件の場合のみ正式出力
tax-system export 2 --template-id 2 --output "outputs/相対表.xlsx"
```

期末在庫表の取込・仕入対応付け・内訳復元は、現時点ではWeb画面からのみ操作できます。

## セキュリティ

- `.tax-system/` には実データを含むため、クラウド同期やGit管理の対象にしないでください。
- Gitへ追加する前に `git status` と `git check-ignore <ファイル>` を確認してください。
- 確認用Excelには「確認用（正式帳票ではありません）」と表示されます。
- 出力履歴にはファイル名、ハッシュ、テンプレート版、検証結果が記録されます。
- 共有用のテストデータは、氏名・住所・電話番号・生年月日・証憑番号・金額を匿名化してください。

## 現在確認できている既存仕様

- 古物台帳CSV: UTF-8 BOM付き、CRLF、11列
- 相対表: `輸出販売` と `免税販売` の2シート
- 相対表は1行目が `受入れ` / `払出し`、2行目が列見出し、3行目以降が明細
- OPTCGとSportCardでは列数・税額計算式が異なるため、テンプレートを別々に版管理
- 期末在庫表CSV: UTF-8 BOM付き、`商品名 / 仕入れ原価 / 在庫数` の3列

詳細は [docs/report-output-spec.md](docs/report-output-spec.md) を参照してください。

## 作り方・技術メモ

- Python + Flask（Web画面）+ SQLite + openpyxl。サーバーは 127.0.0.1 のみで待ち受け。
- 商品名の英語翻訳・内訳復元の組み合わせ提案は OpenAI API（`OPENAI_API_KEY`）を使うが、このPCにはキーを設定していないので、今は手入力に切り替わる。今後、品名の対応表＋プログラム計算／ローカルLLM（Ollama）に置き換える案を検討中。
- 複数PCでGoogleドライブ同期する運用（`TAX_SYSTEM_HOME`）は、このPC1台運用にしたのでやめた。

