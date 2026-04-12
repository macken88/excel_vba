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

## 重要事項

- VBA の編集対象は workbook バイナリではなく `src/` 配下のテキストソースです。
- 同期手順は Codex skill を使って標準化してください。
- `config/*.toml` は実際の workbook 名とソースディレクトリに合わせて更新してください。
- `config/*.toml` はこのリポジトリ独自の `[project]` 形式です。`excel-vba --config` に直接渡す CLI 用設定としては扱わず、skill 用のプロジェクト設定として扱ってください。
- import / export の前に、Excel の `Trust access to the VBA project object model` を有効にしてください。
- import / export skill は `config/*.toml` を読み取り、`excel-vba --file ... --vba-directory ...` の明示引数へ変換して実行してください。
