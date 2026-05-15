# まとめーる / Matomail

**まとめーる** は，Gmail の直近メールを対話的に確認し，LLM による要約・重要度判定・返信草案作成・会議予定の抽出を支援するローカル実行型メール確認アプリです．

本アプリは，メール処理を完全自動化することではなく，ユーザーと対話しながら「読むべきメール」「返信すべきメール」「予定登録すべきメール」を効率よく整理することを目的とします．

## 目的

大学教員・研究者・管理業務担当者のように，多数のメールを日常的に処理するユーザーを想定し，以下を支援します．

- Gmail の直近メールを取得する
- 一度処理したメールを再処理しない
- メール本文全文を LLM に渡して要約する
- 重要度，返信要否，期限，会議予定などを抽出する
- 対話的にユーザーへ判断を求める
- 必要に応じて返信草案を作成する
- ユーザー確認後に Gmail から送信できるようにする
- 会議日程などを Google Calendar に登録する
- 処理結果を HTML レポートとして保存し，ブラウザで確認できるようにする

## 基本方針

まとめーるは，以下の方針で設計します．

1. **Gmail 専用**
   - 初期版では Gmail API のみに対応します．
   - IMAP や大学独自メールサーバーへの対応は対象外です．

2. **手動実行**
   - ユーザーが明示的に起動したときのみメールを確認します．
   - 常駐監視や自動定期実行は初期版では行いません．

3. **最大過去 1 週間**
   - 取得対象は実行時点から過去 7 日以内のメールに限定します．

4. **既読・未読ではなく処理済み管理**
   - Gmail 上の既読／未読とは別に，アプリ側で「処理済みメール」を管理します．
   - 一度処理した Gmail message ID は次回以降スキップします．

5. **添付ファイルは開かない**
   - 添付ファイルは存在のみ検出します．
   - ユーザーが明示的に許可するまで，添付ファイルの本文解析・ダウンロード・LLM 送信は行いません．

6. **送信はユーザー確認後**
   - LLM は返信草案を作成できます．
   - ただし，送信はアプリ上でユーザーが明示的に実行した場合のみ行います．

7. **予定登録は確認付きで実行**
   - 会議や締切などの候補は自動抽出します．
   - Google Calendar への登録は，ユーザー確認後に実行します．

## 想定ワークフロー

1. ユーザーがアプリを起動する
2. Gmail API で過去 1 週間の未処理メールを取得する
3. 各メールについて，本文全文を LLM に渡して解析する
4. アプリがメールごとに以下を提示する
   - 要約
   - 重要度
   - 返信要否
   - 推奨アクション
   - 会議・締切・予定候補
   - 添付ファイルの有無
5. ユーザーが対話的に判断する
   - スキップ
   - 処理済みにする
   - 返信草案を作る
   - 草案を修正する
   - Gmail から送信する
   - カレンダーに登録する
   - 添付ファイルを開くことを許可する
6. 処理結果を HTML にまとめる
7. 指定フォルダに HTML を保存する
8. ブラウザで HTML レポートを表示する

## 出力される HTML レポート

実行ごとに，以下のような HTML レポートを生成します．

```text
reports/
  2026-05-14_1530_matomail.html
```

HTML には以下を含めます．

- 実行日時
- 処理対象メール数
- スキップ済みメール数
- 重要メール一覧
- 要返信メール一覧
- 会議・予定候補一覧
- 各メールの要約
- 推奨アクション
- 作成済み返信草案
- カレンダー登録状況
- 添付ファイルの有無
- Gmail で開くリンク

## 主な機能

### Gmail 取得

- Gmail API を使用してメールを取得する
- 対象期間は最大過去 7 日
- 処理済み Gmail message ID はローカル DB に保存し，再処理を防ぐ
- スレッド情報を取得し，必要に応じて会話文脈を含める

### LLM 解析

各メールに対して，LLM は以下を JSON 形式で返します．

- `summary_ja`
- `category`
- `priority`
- `requires_reply`
- `suggested_action_ja`
- `deadline_candidates`
- `meeting_candidates`
- `attachment_action_required`
- `reply_recommended`
- `reply_draft_ja`
- `confidence`

### 対話モード

CLI または Web UI 上で，メールごとに以下を尋ねます．

- このメールを処理済みにしますか？
- 返信草案を作成しますか？
- 草案を修正しますか？
- この草案を Gmail から送信しますか？
- 会議候補を Google Calendar に登録しますか？
- 添付ファイルを開いて解析してよいですか？
- このメールを後回しにしますか？

### Google Calendar 連携

- メール本文から会議候補を抽出する
- 日時，参加者，件名，場所，オンライン会議 URL を推定する
- ユーザー確認後に Google Calendar へ登録する
- 不確実な場合は自動登録せず，確認待ちにする

### Gmail 送信

- 返信草案を作成する
- ユーザーが修正できる
- ユーザーが明示的に送信ボタンを押した場合のみ Gmail API 経由で送信する

## 技術スタック案

初期実装では以下を想定します．

- Python 3.11+
- Gmail API
- Google Calendar API
- SQLite
- SQLAlchemy
- Jinja2
- FastAPI または Streamlit
- OpenAI API または任意の LLM API
- pytest

## 開発環境のセットアップ

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

設定は `.env.example` をもとに `.env` を作成します．

Gmail を実際に読み込む場合は，Google Cloud で OAuth クライアントを作成し，ダウンロードした JSON を `credentials.json` としてリポジトリ直下に置きます．`credentials.json` と認証後に生成される `token.json` は Git 管理対象外です．

OAuth クライアントは Desktop app として作成するのが簡単です．Web application として作成した場合は，Google Cloud Console の「承認済みのリダイレクト URI」に以下を追加してください．

```text
http://localhost:8080/
```

```powershell
.\.venv\Scripts\matomail.exe fetch
```

初回実行時はブラウザで Google の認証フローが開きます．現在の実装では Gmail の読み取り専用スコープを使い，添付ファイルはメタデータのみ検出してダウンロードしません．

`Gmail API has not been used ... or it is disabled` と表示された場合は，Google Cloud Console で Gmail API を有効化してから再実行してください．有効化直後は反映まで数分かかることがあります．

取得したメールのメタデータ，本文全文，添付メタデータ，処理状態は `MATOMAIL_DB_PATH` に指定したメール用 SQLite DB に保存されます．初期値は `./data/matomail.sqlite3` です．同じ Gmail message ID のメールは重複保存されず，`processed` または `skipped` にしたメールは次回以降の処理対象から除外されます．`pending` のメールは後で再確認できる状態として残ります．

本文保存は `MATOMAIL_STORE_EMAIL_BODY=false` で無効化できます．DB が `MATOMAIL_DB_MAX_SIZE_MB` を超えた場合は，起動時に `MATOMAIL_DB_BACKUP_DIR` 配下の日時付きフォルダへ既存 DB を退避し，新しい `matomail.sqlite3` から再開します．

恒久的な判定条件は `MATOMAIL_RULES_DB_PATH` に指定したルール用 SQLite DB に保存します．初期値は `./data/matomail_rules.sqlite3` です．Gmail のフィルタ条件を参考に，送信者，宛先，件名，含む語句，含まない語句，添付有無，添付ファイル名，サイズ，受信日時などを持てる設計です．アクションは `ignore`，`always_process`，`skip_analysis` を想定します．Gmail ラベルやカテゴリなど，取得側のメタデータがまだない項目は DB に保存できる形だけ先に用意しています．

単純なフィルタでは表現しにくい判断基準は `llm_instruction_rules` テーブルに保存します．たとえば「査読対応の依頼は Abstract を読んで専門との合致率を表示する」「本文中に特定語がない場合は優先度を下げる」「学位プログラムの Web 更新は優先し，学類の Web 更新は低優先度にする」といった追加指示を，常時または条件付きで LLM プロンプトへ差し込めます．

fetch 時にはルール DB を参照し，判定結果をメール DB の `filter_decisions` テーブルに保存します．`preclassify` ルールに一致したメールは，LLM を使わずに `email_analysis` へ自動分類結果を保存できます．

LLM 解析には OpenAI API を使います．API キーは `.env` の `OPENAI_API_KEY` に設定し，利用モデルは `MATOMAIL_LLM_MODEL` で指定します．`.env` は Git 管理対象外です．

ルールは CLI から管理できます．

```powershell
.\.venv\Scripts\matomail.exe rules list
.\.venv\Scripts\matomail.exe rules add-sender-skip newsletter@example.com --name "newsletter"
.\.venv\Scripts\matomail.exe rules add-subject-skip "Dr. Horie - Read the most recent articles online"
.\.venv\Scripts\matomail.exe rules add-subject-preclassify "緊急" --priority high --category important --summary "緊急メール"
.\.venv\Scripts\matomail.exe rules disable 1
.\.venv\Scripts\matomail.exe rules enable 1
.\.venv\Scripts\matomail.exe rules delete 1
```

LLM 追加指示も CLI から管理できます．

```powershell
.\.venv\Scripts\matomail.exe instructions list
.\.venv\Scripts\matomail.exe instructions add "査読対応の依頼は Abstract を読み，専門との合致率を表示する。" --subject-query "査読"
.\.venv\Scripts\matomail.exe instructions disable 1
.\.venv\Scripts\matomail.exe instructions enable 1
.\.venv\Scripts\matomail.exe instructions delete 1
```

## ディレクトリ構成案

```text
matomail/
  README.md
  REQUIREMENTS.md
  pyproject.toml
  .env.example
  src/
    matomail/
      __init__.py
      app.py
      config.py
      gmail_client.py
      calendar_client.py
      llm_client.py
      analyzer.py
      interaction.py
      report.py
      database.py
      models.py
      prompts/
        email_analysis.md
        reply_draft.md
      templates/
        report.html.j2
  data/
    matomail.sqlite3
  reports/
  tests/
    test_gmail_client.py
    test_analyzer.py
    test_report.py
```

## セキュリティとプライバシー

- メール本文全文を LLM に送るため，利用する LLM API のデータ保持ポリシーを確認する
- 添付ファイルは明示的な許可なしに開かない
- OAuth トークンは Git 管理しない
- `.env`，`token.json`，`credentials.json`，SQLite DB，HTML レポートは `.gitignore` に含める
- HTML レポートには個人情報が含まれる可能性があるため，公開リポジトリにはアップロードしない

## 初期版でやらないこと

- メールの常時監視
- メールの自動削除
- 添付ファイルの自動解析
- ユーザー確認なしの返信送信
- ユーザー確認なしのカレンダー登録
- Gmail 以外のメールサービス対応
- 複数ユーザー運用
- サーバー公開運用


個人利用・研究室内利用を想定する場合でも，メール本文や認証情報を含むデータはリポジトリに含めないでください．
