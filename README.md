# VBA 保守テンプレート

このリポジトリテンプレートは、VBA を含む Excel ブックを Git 管理で保守するためのものです。

## 構成

- `workbooks/`: `.xlsm` などの Excel ブック本体
- `src/BookA_vba/`: BookA から書き出した VBA ソース
- `src/BookB_vba/`: BookB から書き出した VBA ソース
- `src/shared/`: 複数ブックで共有する VBA モジュール
- `config/`: Codex skill が参照する、ブックとソースの対応設定
- `AGENTS.md`: Codex と人間の保守担当者向けの運用ルール
- `requirements-vba-tools.txt`: 固定した Python ツール依存関係

## セットアップ例

Windows 上で専用の仮想環境を作成します。

```powershell
python -m venv .venv-vba-tools
.\\.venv-vba-tools\\Scripts\\python.exe -m pip install --upgrade pip
.\\.venv-vba-tools\\Scripts\\python.exe -m pip install -r requirements-vba-tools.txt
```

## セットアップ後の確認

環境が正しく構築されたことを以下のコマンドで確認できます。

```powershell
# vba-edit がインストールされていることを確認
.\\.venv-vba-tools\\Scripts\\python.exe -m pip show vba-edit

# config/*.toml で指定されたブックが存在するか確認
Get-ChildItem config\\*.toml | ForEach-Object {
    $file = Select-String -LiteralPath $_.FullName -Pattern '^file\s*=\s*"(.+)"$' |
        Select-Object -First 1
    if ($file) {
        $path = $file.Matches[0].Groups[1].Value
        Write-Host "$($_.Name): $path -> exists=$(Test-Path $path)"
    }
}
```

## 重要事項

- VBA の編集対象は workbook バイナリではなく `src/` 配下のテキストソースです。
- 同期手順は Codex skill を使って標準化してください。
- `config/*.toml` は実際の workbook 名とソースディレクトリに合わせて更新してください。
- `config/*.toml` はこのリポジトリ独自の `[project]` 形式です。`excel-vba --config` に直接渡す CLI 用設定としては扱わず、skill 用のプロジェクト設定として扱ってください。
- import / export の前に、Excel の `Trust access to the VBA project object model` を有効にしてください。
- import / export skill は `config/*.toml` を読み取り、`excel-vba --file ... --vba-directory ...` の明示引数へ変換して実行してください。

## Git ワークフロー

### ブランチ戦略

main ブランチは常に Excel へ import 可能な安定状態を保ちます。変更は作業ブランチで行い、レビュー後に main へマージします。

```
main                 ← 常に安定（import 可能）
 └── feature/...     ← 作業ブランチ（短命）
```

#### ブランチ命名規則

| パターン | 用途 | 例 |
|---|---|---|
| `feature/<book>/<説明>` | ブック固有の機能追加 | `feature/bookA/add-validation` |
| `fix/<book>/<説明>` | ブック固有の不具合修正 | `fix/bookB/fix-date-parse` |
| `shared/<説明>` | 共有モジュールの変更 | `shared/update-common-utilities` |
| `chore/<説明>` | ツール・設定変更 | `chore/update-vba-edit` |
| `docs/<説明>` | ドキュメントのみの変更 | `docs/add-git-workflow-rules` |

- `.xlsm` バイナリは Git でマージできないため、ブランチは短命に保つこと
- マージ済みブランチは再利用しない。バグ修正は main から新しい `fix/` ブランチを切る

### コミットメッセージ

[Conventional Commits](https://www.conventionalcommits.org/) 形式を使用します。

```
<type>(<scope>): <summary>

[body]
```

#### type 一覧

| type | 意味 | 例 |
|---|---|---|
| `feat` | VBA 機能の追加・変更 | `feat(BookA): マクロに入力チェックを追加` |
| `fix` | VBA の不具合修正 | `fix(shared): CommonUtilities の日付処理を修正` |
| `sync` | Excel ブックへの import / export | `sync(BookA): src → xlsm へ import 反映` |
| `refactor` | 動作を変えないコード整理 | `refactor(BookB): 変数名を統一` |
| `chore` | ツール・設定変更 | `chore: vba-edit を 0.3.1 に更新` |
| `docs` | ドキュメントのみの変更 | `docs: AGENTS.md にブランチ規約を追加` |

#### scope ルール

- ブック固有の変更: ブック名（`BookA`, `BookB`）
- 共有モジュール: `shared`
- 設定・ツール: 省略可

#### body

- Excel への re-import の要否を必ず明記する（例: `Re-import required: BookA`）
- 共有モジュール変更の場合: 影響を受けるブック一覧を記載する

### コミット分離の原則

- `src/` のテキスト変更と `.xlsm` バイナリ同期は別コミットにする
- `src/shared/` の変更と、各ブックディレクトリへのコピーは別コミットにする
- 複数ブックへの変更は、ブック単位でコミットを分離することを推奨する

### プルリクエスト

- 1 つの改修（機能追加・変更・バグ修正、またはそれらの組み合わせ）を 1 PR とする
- 改修内の個々の変更はコミット単位で分離する（目的別・ブック別・sync 別）
- 無関係なブックへの変更や、改修と無関係なリファクタリングは別 PR にする
- PR の説明には、影響するブック名と re-import 要否を必ず記載する
- 元の改修で生じたバグ修正は新しい PR とし、元の PR 番号を参照して関連を示す
