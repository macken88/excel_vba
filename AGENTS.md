# AGENTS.md

## 目的
このリポジトリは、VBA を含む Excel ブックを長期的に保守・メンテナンスするためのものです。  
主な目的は次のとおりです。

- VBA ソースコードを Git で管理すること
- AI エージェントによる安全な補助編集を可能にすること
- Excel ブック本体と、書き出した VBA テキストソースを、管理された手順で同期すること

## 対象範囲
- 保守対象の Excel ブックは `workbooks/` に配置する
- Git 管理対象の VBA ソースは `src/` に配置する
- 複数ブックで共有する共通モジュールは `src/shared/` に配置する
- ブックとソースの対応を表すリポジトリ用設定ファイルは `config/` に配置する
- プロジェクトで定義する Codex skill は `.agents/skills/` に配置する

## 基本方針
1. 原則として、Excel の VBA エディタ上で直接コード編集しない
2. Git 管理対象の正本は `src/` 配下の `.bas` / `.cls` / `.frm` とする
3. `.xlsm` などのバイナリブックは、通常利用および管理された同期手順以外では変更しない
4. 変更は、追跡可能でレビューしやすいテキスト差分として行う
5. Excel への反映前には必ず Git diff を確認する

## AI エージェント向けルール
1. `.xlsm` ファイルを直接編集しない
2. 編集対象は `src/` 配下のテキストベース VBA ソースに限定する
3. セットアップ、export、edit、共有モジュール同期、import などの定型作業は、必ずプロジェクトで定義された Codex skill を経由して実行する
4. 既存 skill で対応できる作業について、新しい同期手順や独自手順を勝手に作らない
5. 共有モジュールを変更する場合は、まず `src/shared/` を更新し、その後で各ブック用ディレクトリへ反映する
6. 明示的な依頼がない限り、変更は最小限かつ局所的に行う
7. モジュール名やブック固有の構成は、明示的な依頼がない限り変更しない
8. `config/*.toml` のブック対象やディレクトリ設定を、理由なく変更しない
9. `.frm` や class header まわりの不整合が疑われる場合は、推測で直さず、export をやり直す前提で対処する
10. 作業後は、変更したファイルと、Excel 側への再 import が必要かどうかを必ず要約する
11. `config/*.toml` はこのリポジトリ独自の `[project]` 形式として扱い、`vba-edit` の `--config` にそのまま渡せる前提で扱わない
12. import / export 系 skill は、`config/*.toml` から `file` と `vba_directory` を解釈して `vba-edit` CLI の明示引数へ変換する
13. 既存 skill の参照先は `.agents/skills/` を正本とし、`.codex/skills/` は使わない
14. 新しい skill を追加する場合も、作成先は `.agents/skills/` に統一する

## `config/*.toml` のスキーマ

各 TOML ファイルはブックごとに 1 つ作成し、ファイル名はブック名に対応させる（例: `bookA.toml` → `workbooks/BookA.xlsm`）。

### `[project]` セクション

| キー | 型 | 必須 | 説明 |
|---|---|---|---|
| `file` | string | はい | 対象 Excel ブックのリポジトリルートからの相対パス（例: `"workbooks/BookA.xlsm"`） |
| `vba_directory` | string | はい | VBA ソースの格納ディレクトリのリポジトリルートからの相対パス（例: `"src/BookA_vba"`） |
| `encoding` | string | いいえ | ソースファイルのエンコーディング。将来の拡張用。現在のスクリプトでは使用していない |

### `[headers]` セクション

| キー | 型 | 必須 | 説明 |
|---|---|---|---|
| `save_headers` | boolean | いいえ | VBA export 時にモジュールヘッダ（`Attribute` 行など）を保存するか。将来の拡張用。現在のスクリプトでは使用していない |
| `in_file_headers` | boolean | いいえ | ソースファイル内にヘッダ情報を含めるか。将来の拡張用。現在のスクリプトでは使用していない |

### ブック追加手順

新しい Excel ブックを管理対象に追加する場合は、以下の手順を行う。

1. `.xlsm` ブックを `workbooks/` に配置する
2. `src/` 配下にブック用の VBA ソースディレクトリを作成する（例: `src/BookC_vba/`）
3. `config/` にブック名に対応する TOML ファイルを作成する（例: `config/bookC.toml`）
4. TOML ファイルに `[project]` セクションの `file` と `vba_directory` を記述する
5. 必要に応じて `export-vba` skill で既存ブックから VBA ソースを書き出す

## 環境に関するルール
1. このリポジトリは Windows 上で運用することを前提とする  
   理由: Excel / Office automation が Windows 前提のため
2. VBA 保守ツールは専用の Python 仮想環境で管理する
3. ツール依存関係は `requirements-vba-tools.txt` からインストールする
4. `vba-edit` のバージョンは固定し、明示的な承認なしに更新しない

## Codex に期待する skill
Codex には、最終的に次のような skill を用意することを想定する。

- `export-vba`
- `edit-vba`
- `sync-shared-modules`
- `import-vba`
- `review-vba-diff`

これらの skill は、場当たり的なシェル操作ではなく、再現可能で決定的な手順をカプセル化すること。
配置先は `.agents/skills/` を標準とする。

## Git 運用ルール
1. commit 前に必ず diff を確認する
2. commit は小さく、意味のある単位で分ける
3. 共有モジュールの配布と、Excel への import は、可能な限り別ステップで扱う
4. バイナリの Excel ブックが想定外に変更されていた場合は、その事実を明示する

## 注意事項
- `vba-edit` は、Office の VBA Project とテキストファイルを同期するための Windows 向けツールである
- `config/*.toml` はリポジトリ運用用の設定であり、`vba-edit` ネイティブの config 仕様とは一致しない場合がある
- Excel 側の信頼設定により同期できない場合は、import / export skill 内の Trust Access チェックで検出される。ガードレールを迂回しない
- import / export 前には `Trust access to the VBA project object model` が有効であることを確認する
- 直接 Excel 側を触って解決するよりも、`src/` を正本とする運用を優先する
