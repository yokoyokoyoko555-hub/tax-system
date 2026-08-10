# tax-system

古物台帳と輸出・免税販売の仕入販売相対表を、既存フォーマットのまま管理・検証・再出力するローカルシステムです。

実データ、テンプレートの複製、SQLiteデータベース、出力ファイルは `.tax-system/` または `outputs/` に保存され、Git管理から除外されます。原本は読み取り専用で参照し、直接更新しません。

## セットアップ

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e .
tax-system init
```

## 基本操作

```powershell
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

## 複数PCでの利用（Googleドライブ同期）

Googleドライブの同期フォルダをデータ保存先にすることで、複数のPCから同じデータを扱えます。

```powershell
# 各PCで一度だけ設定（PowerShellのプロファイルに追記すると毎回設定不要）
$env:TAX_SYSTEM_HOME = "G:\マイドライブ\tax-system-data"

# 以降は --home を省略しても環境変数の場所が使われる
tax-system init
```

**注意（SQLiteの制約）**

- 同時に2台のPCから開かないこと。片方で操作後、Googleドライブの同期が完了してから、もう片方のPCを開いてください。同時に書き込むとDBが壊れたり、`.sqlite3 (競合するコピー)` のような同期エラーファイルが作られる場合があります。
- 実データ（氏名・住所・電話番号等）がGoogleドライブ上に置かれます。共有設定はリンク共有をオフにし、自分のアカウントのみアクセスできる状態にしてください。
- 出力先（`--output`）にもGoogleドライブ上のパスを指定すれば、`outputs/` を共有できます。

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

詳細は [docs/report-output-spec.md](docs/report-output-spec.md) を参照してください。

