# まとめーる 要件定義書

## 1. 概要

### 1.1 名称

まとめーる / Matomail

### 1.2 目的

まとめーるは，Gmail の直近メールを LLM により要約・分類し，ユーザーとの対話を通じて返信，予定登録，添付確認，処理済み管理を支援するローカル実行型アプリケーションである．

本システムはメール処理の完全自動化を目的としない．ユーザーの判断を維持しつつ，メール確認に伴う認知負荷と作業時間を削減することを目的とする．

### 1.3 想定ユーザー

- 大学教員
- 研究者
- 研究室管理者
- 共同研究・学生指導・事務連絡を多数処理するユーザー

### 1.4 初期スコープ

- Gmail 専用
- 手動実行
- 最大過去 1 週間のメール確認
- 本文全文を LLM に渡して解析
- 添付ファイルはユーザー許可まで開かない
- 返信草案作成
- ユーザー確認後の Gmail 送信
- 会議予定候補の抽出
- ユーザー確認後の Google Calendar 登録
- HTML レポート生成
- 処理済みメールのスキップ

## 2. 機能要件

### 2.1 Gmail 認証

#### REQ-GMAIL-001

システムは Gmail API を利用してユーザーの Gmail アカウントに OAuth 認証できること．

#### REQ-GMAIL-002

OAuth 認証情報はローカルに保存し，Git 管理対象から除外すること．

#### REQ-GMAIL-003

初回実行時に認証が完了していない場合，ユーザーに認証フローを案内すること．

### 2.2 メール取得

#### REQ-MAIL-001

システムは実行時点から過去 7 日以内の Gmail メールを取得できること．

#### REQ-MAIL-002

取得対象は，Gmail message ID，thread ID，送信者，宛先，CC，件名，受信日時，本文，スニペット，添付ファイル情報を含むこと．

#### REQ-MAIL-003

一度処理済みと記録された Gmail message ID は，次回以降の処理対象から除外すること．

#### REQ-MAIL-004

Gmail 上の既読・未読状態とは独立して，アプリ側で処理済み状態を管理すること．

#### REQ-MAIL-005

取得件数が多い場合，ユーザーに処理上限を指定できることが望ましい．

### 2.3 添付ファイル処理

#### REQ-ATTACH-001

システムは添付ファイルの有無，ファイル名，MIME type，サイズを取得できること．

#### REQ-ATTACH-002

システムはユーザーが明示的に許可するまで，添付ファイルをダウンロードまたは解析してはならない．

#### REQ-ATTACH-003

添付ファイルがある場合，メール要約画面で「添付あり」と明示すること．

#### REQ-ATTACH-004

ユーザーが許可した場合のみ，添付ファイルを取得し，必要に応じて LLM 解析対象に含めること．

### 2.4 LLM 解析

#### REQ-LLM-001

システムはメール本文全文を LLM に渡し，日本語で要約を生成できること．

#### REQ-LLM-002

システムは LLM により以下の項目を抽出できること．

- 要約
- カテゴリ
- 重要度
- 返信要否
- 推奨アクション
- 締切候補
- 会議候補
- 予定登録候補
- 添付ファイル確認要否
- 返信草案の要否
- 返信草案
- 解析信頼度

#### REQ-LLM-003

LLM の出力は原則 JSON 形式とし，パース可能であること．

#### REQ-LLM-004

LLM 出力が不正な JSON の場合，再試行またはフォールバック処理を行うこと．

#### REQ-LLM-005

重要度は少なくとも `high`, `medium`, `low` の 3 段階で表現すること．

#### REQ-LLM-006

返信草案は，送信者，文脈，ユーザーの立場を踏まえた自然な文面で作成すること．

#### REQ-LLM-007

不確実な日時，曖昧な依頼，判断が必要な内容については，自動実行せずユーザー確認対象とすること．

### 2.5 対話処理

#### REQ-INTERACT-001

システムはメールごとにユーザーへ処理方針を確認できること．

#### REQ-INTERACT-002

ユーザーは各メールに対して以下の操作を選択できること．

- 処理済みにする
- 後で確認する
- 返信草案を作成する
- 返信草案を修正する
- 返信を送信する
- カレンダーに登録する
- 添付ファイルを開く
- スキップする

#### REQ-INTERACT-003

システムは LLM の推奨アクションを提示したうえで，ユーザーの判断を求めること．

#### REQ-INTERACT-004

ユーザーが明示的に承認しない限り，メール送信，カレンダー登録，添付ファイル解析は実行してはならない．

### 2.6 返信草案作成・送信

#### REQ-REPLY-001

システムは LLM により返信草案を作成できること．

#### REQ-REPLY-002

返信草案はアプリ上で編集可能であること．

#### REQ-REPLY-003

ユーザーが送信を承認した場合のみ，Gmail API を用いて返信を送信できること．

#### REQ-REPLY-004

返信は可能な限り元メールのスレッドに紐づけること．

#### REQ-REPLY-005

送信済み返信の本文，送信日時，対象 message ID を処理ログに記録すること．

### 2.7 Google Calendar 連携

#### REQ-CAL-001

システムはメール本文から会議候補を抽出できること．

#### REQ-CAL-002

会議候補には以下を含むこと．

- 件名
- 開始日時
- 終了日時
- タイムゾーン
- 場所
- オンライン会議 URL
- 参加者候補
- 説明文
- 抽出信頼度

#### REQ-CAL-003

日時が曖昧な場合は，自動登録せず確認待ちとすること．

#### REQ-CAL-004

ユーザーが承認した場合のみ，Google Calendar API により予定を登録すること．

#### REQ-CAL-005

登録済み予定の calendar event ID を処理ログに記録すること．

### 2.8 HTML レポート生成

#### REQ-REPORT-001

システムは実行ごとに HTML レポートを生成すること．

#### REQ-REPORT-002

HTML レポートは指定フォルダに保存すること．

#### REQ-REPORT-003

HTML レポート生成後，既定ブラウザで開けること．

#### REQ-REPORT-004

HTML レポートには以下を含むこと．

- 実行日時
- 取得メール数
- 処理済みにしたメール数
- スキップしたメール数
- 重要メール一覧
- 要返信メール一覧
- 会議候補一覧
- 各メールの要約
- 推奨アクション
- 返信草案
- 送信状況
- カレンダー登録状況
- 添付ファイル情報
- Gmail で開くリンク

#### REQ-REPORT-005

HTML レポートはローカル閲覧を想定し，外部公開を前提としないこと．

### 2.9 処理済み管理

#### REQ-STATE-001

システムは処理済み Gmail message ID を SQLite などのローカル DB に保存すること．

#### REQ-STATE-002

処理済み状態には以下を含めること．

- message ID
- thread ID
- 処理日時
- 処理種別
- 返信草案作成有無
- 返信送信有無
- カレンダー登録有無
- 添付解析有無
- レポート出力先

#### REQ-STATE-003

ユーザーが「後で確認」を選んだメールは，処理済みではなく保留状態として記録すること．

#### REQ-STATE-004

保留状態のメールは次回実行時に再提示できること．

### 2.10 フィルタリング管理

#### REQ-FILTER-001

システムは，恒久的に解析不要とする条件を SQLite などのローカル DB に保存できること．

フィルタルールは，メール本文・解析結果・処理状態を保存するメール DB とは別のルール DB に保存すること．

#### REQ-FILTER-002

初期実装では，送信者メールアドレス単位で「常に解析不要」とするルールを登録できること．

#### REQ-FILTER-003

フィルタルールには，名前，アクション，優先度，Gmail 風の検索条件，有効／無効，メモ，作成日時を含めること．

#### REQ-FILTER-004

フィルタに一致したメールは，LLM 解析対象から除外，強制解析，または LLM を使わない自動分類の対象にできること．

#### REQ-FILTER-005

fetch 時に実行したフィルタ判定結果は，メール DB に保存すること．判定結果にはアクション，マッチしたルール ID，ルール名，理由，判定時点のルールスナップショット，判定日時を含めること．

#### REQ-FILTER-006

`preclassify` ルールに一致したメールは，LLM 解析を実行せず，ルールに定義された優先度，カテゴリ，要約，推奨アクションなどを `email_analysis` に保存できること．

### 2.11 LLM 追加指示管理

#### REQ-INSTRUCTION-001

システムは，LLM 解析時に追加で適用するユーザー指示を SQLite などのローカル DB に保存できること．

LLM 追加指示は，メール本文・解析結果・処理状態を保存するメール DB とは別のルール DB に保存すること．

#### REQ-INSTRUCTION-002

LLM 追加指示は，常時適用または Gmail 風の条件に一致したメールのみに適用できること．

#### REQ-INSTRUCTION-003

LLM 追加指示には，名前，指示本文，優先度，条件，有効／無効，メモ，作成日時を含めること．

#### REQ-INSTRUCTION-004

条件に一致した LLM 追加指示は，メール解析プロンプトへ優先度順に差し込めること．

#### REQ-RULECLI-001

システムは CLI からフィルタルールを一覧表示，追加，有効化，無効化，削除できること．

#### REQ-RULECLI-002

システムは CLI から LLM 追加指示ルールを一覧表示，追加，有効化，無効化，削除できること．

## 3. 非機能要件

### 3.1 セキュリティ

#### NFR-SEC-001

OAuth 認証情報，API key，LLM key は Git 管理しないこと．

#### NFR-SEC-002

`.env`，`credentials.json`，`token.json`，SQLite DB，HTML レポートは `.gitignore` に含めること．

#### NFR-SEC-003

HTML レポートには個人情報やメール本文要約が含まれるため，公開リポジトリに含めないこと．

#### NFR-SEC-004

添付ファイルはユーザー承認なしに解析しないこと．

#### NFR-SEC-005

メール送信，予定登録，添付解析は必ずユーザー承認を必要とすること．

### 3.2 プライバシー

#### NFR-PRIV-001

メール本文全文を LLM に送信する設計であるため，利用する LLM API のデータ利用ポリシーを確認できるよう README に明記すること．

#### NFR-PRIV-002

可能であれば，ログにはメール本文全文を保存せず，要約と処理結果のみを保存すること．

#### NFR-PRIV-003

デバッグログに本文全文，添付ファイル本文，認証情報を出力しないこと．

### 3.3 可用性

#### NFR-AVL-001

Gmail API，Google Calendar API，LLM API のいずれかでエラーが発生しても，処理済み状態が破損しないこと．

#### NFR-AVL-002

途中で処理が中断された場合でも，完了済みメールと未完了メールを区別できること．

### 3.4 保守性

#### NFR-MAINT-001

Gmail 取得，LLM 解析，対話処理，レポート生成，カレンダー登録はモジュールを分離すること．

#### NFR-MAINT-002

LLM プロンプトはコードから分離し，`prompts/` 配下で管理すること．

#### NFR-MAINT-003

主要処理には pytest によるテストを用意すること．

## 4. データ設計案

### 4.1 emails

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| gmail_message_id | text | Gmail message ID |
| gmail_thread_id | text | Gmail thread ID |
| sender | text | sender |
| recipients | text | recipients JSON |
| subject | text | subject |
| received_at | datetime | received timestamp |
| snippet | text | Gmail snippet |
| has_attachments | boolean | attachment flag |
| created_at | datetime | record creation timestamp |

### 4.2 email_analysis

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| email_id | integer | emails.id |
| summary_ja | text | Japanese summary |
| category | text | category |
| priority | text | high / medium / low |
| requires_reply | boolean | reply required |
| suggested_action_ja | text | suggested action |
| deadline_candidates_json | text | deadline candidates |
| meeting_candidates_json | text | meeting candidates |
| reply_draft | text | draft reply |
| confidence | real | confidence |
| llm_model | text | model name |
| created_at | datetime | analysis timestamp |

### 4.3 processing_state

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| gmail_message_id | text | Gmail message ID |
| status | text | processed / pending / skipped |
| action_taken | text | summary of action |
| reply_sent | boolean | reply sent |
| calendar_registered | boolean | calendar event registered |
| attachment_opened | boolean | attachment opened |
| report_path | text | generated report path |
| updated_at | datetime | timestamp |

### 4.4 filter_decisions

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| email_id | integer | emails.id |
| gmail_message_id | text | Gmail message ID |
| action | text | ignore / always_process / skip_analysis / preclassify |
| matched_rule_id | integer | rule ID in rules DB |
| matched_rule_name | text | rule name snapshot |
| reason | text | decision reason |
| rule_snapshot_json | text | rule snapshot JSON |
| decided_at | datetime | decision timestamp |

### 4.5 filter_rules

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| name | text | rule name |
| action | text | ignore / always_process / skip_analysis / preclassify |
| priority | integer | higher priority wins |
| preset_priority | text | preclassified priority |
| preset_category | text | preclassified category |
| preset_summary_ja | text | preclassified summary |
| preset_suggested_action_ja | text | preclassified suggested action |
| preset_requires_reply | boolean | preclassified reply flag |
| preset_reply_recommended | boolean | preclassified reply recommendation |
| from_query | text | sender condition |
| to_query | text | recipient condition |
| delivered_to_query | text | delivered-to condition |
| cc_query | text | CC condition |
| bcc_query | text | BCC condition, stored for future use |
| subject_query | text | subject contains condition |
| has_words | text | words/query that must be present |
| doesnt_have | text | words/query that must be absent |
| gmail_query | text | Gmail search query, stored for future use |
| negated_gmail_query | text | negated Gmail search query, stored for future use |
| has_attachment | boolean | attachment condition |
| filename_query | text | attachment filename condition |
| size_comparison | text | larger / smaller |
| size_bytes | integer | message size threshold |
| after | datetime | received after condition |
| before | datetime | received before condition |
| older_than | text | Gmail-style age condition, stored for future use |
| newer_than | text | Gmail-style age condition, stored for future use |
| category | text | Gmail category, stored for future use |
| label | text | Gmail label, stored for future use |
| include_chats | boolean | Gmail-compatible flag |
| note | text | user note |
| enabled | boolean | enabled flag |
| created_at | datetime | creation timestamp |

### 4.6 llm_instruction_rules

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| name | text | rule name |
| instruction | text | additional LLM instruction |
| priority | integer | higher priority first |
| from_query | text | sender condition |
| to_query | text | recipient condition |
| delivered_to_query | text | delivered-to condition |
| cc_query | text | CC condition |
| bcc_query | text | BCC condition, stored for future use |
| subject_query | text | subject contains condition |
| has_words | text | words/query that must be present |
| doesnt_have | text | words/query that must be absent |
| gmail_query | text | Gmail search query, stored for future use |
| negated_gmail_query | text | negated Gmail search query, stored for future use |
| has_attachment | boolean | attachment condition |
| filename_query | text | attachment filename condition |
| size_comparison | text | larger / smaller |
| size_bytes | integer | message size threshold |
| after | datetime | received after condition |
| before | datetime | received before condition |
| older_than | text | Gmail-style age condition, stored for future use |
| newer_than | text | Gmail-style age condition, stored for future use |
| category | text | Gmail category, stored for future use |
| label | text | Gmail label, stored for future use |
| include_chats | boolean | Gmail-compatible flag |
| note | text | user note |
| enabled | boolean | enabled flag |
| created_at | datetime | creation timestamp |

### 4.7 calendar_events

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| email_id | integer | emails.id |
| title | text | event title |
| start_time | datetime | start |
| end_time | datetime | end |
| timezone | text | timezone |
| location | text | location |
| attendees_json | text | attendees |
| calendar_event_id | text | Google Calendar event ID |
| status | text | candidate / registered / rejected |
| created_at | datetime | timestamp |

### 4.8 reply_logs

| Field | Type | Description |
|---|---|---|
| id | integer | internal ID |
| email_id | integer | emails.id |
| draft_body | text | draft body |
| final_body | text | final sent body |
| sent | boolean | sent flag |
| sent_at | datetime | sent timestamp |

## 5. LLM 出力 JSON 案

```json
{
  "summary_ja": "メール内容の要約",
  "category": "学生対応",
  "priority": "high",
  "requires_reply": true,
  "suggested_action_ja": "日程候補を確認し，返信する",
  "deadline_candidates": [
    {
      "date": "2026-05-20",
      "description": "回答期限",
      "confidence": 0.8
    }
  ],
  "meeting_candidates": [
    {
      "title": "研究打ち合わせ",
      "start_time": "2026-05-21T15:00:00+09:00",
      "end_time": "2026-05-21T16:00:00+09:00",
      "timezone": "Asia/Tokyo",
      "location": "Zoom",
      "online_meeting_url": "https://example.com",
      "attendees": ["example@example.com"],
      "confidence": 0.85
    }
  ],
  "attachment_action_required": false,
  "reply_recommended": true,
  "reply_draft_ja": "返信草案本文",
  "confidence": 0.82
}
```

## 6. 画面・インターフェース要件

### 6.1 初期版

初期版では CLI または簡易 Web UI のどちらかで実装する．

推奨は簡易 Web UI である．理由は，返信草案の編集，カレンダー登録候補の確認，添付ファイル確認などが CLI より扱いやすいためである．

### 6.2 メール確認画面

各メールについて以下を表示する．

- 件名
- 送信者
- 受信日時
- 要約
- 重要度
- 返信要否
- 推奨アクション
- 会議候補
- 添付ファイル有無
- 操作ボタン

### 6.3 操作ボタン

- 処理済みにする
- 後で確認
- 返信草案を作る
- 草案を修正
- 送信
- カレンダー登録
- 添付ファイルを開く
- Gmail で開く

## 7. 実装フェーズ

### Phase 1: 最小動作版

- Gmail OAuth
- 過去 1 週間のメール取得
- 処理済み message ID の保存
- LLM 要約
- HTML レポート生成
- ブラウザ表示

### Phase 2: 対話処理

- メールごとの確認 UI
- 処理済み／保留／スキップ管理
- 返信草案生成
- 草案編集

### Phase 3: Gmail 返信送信

- Gmail スレッド返信
- ユーザー確認後の送信
- 送信ログ保存

### Phase 4: Google Calendar 連携

- 会議候補抽出
- カレンダー登録確認 UI
- Google Calendar 登録
- 登録ログ保存

### Phase 5: 添付ファイル許可制解析

- 添付一覧表示
- ユーザー許可後の取得
- PDF / docx / xlsx などの解析拡張

## 8. Codex 向け開発タスク案

### Task 1: Project skeleton

Create a Python project skeleton for Matomail.

Requirements:
- Python 3.11+
- src layout
- pyproject.toml
- pytest
- .env.example
- .gitignore
- README.md
- REQUIREMENTS.md

### Task 2: Gmail OAuth and fetcher

Implement Gmail OAuth and fetch emails from the last 7 days.

Requirements:
- Use Gmail API
- Fetch message metadata and full plain-text body
- Detect attachments but do not download them
- Return structured Python objects
- Add tests with mocked Gmail API responses

### Task 3: SQLite state management

Implement SQLite persistence.

Requirements:
- SQLAlchemy models
- emails table
- email_analysis table
- processing_state table
- filter_rules table
- llm_instruction_rules table
- Avoid duplicate processing by Gmail message ID
- Support pending status
- Support sender-based skip-analysis filters
- Support conditional LLM instruction rules

### Task 4: LLM email analyzer

Implement LLM-based email analyzer.

Requirements:
- Send full email body to LLM
- Use prompt template from prompts/email_analysis.md
- Parse JSON response
- Validate required fields
- Retry once on invalid JSON
- Save analysis to DB

### Task 5: HTML report generator

Generate an HTML report for each run.

Requirements:
- Use Jinja2
- Save to configurable reports directory
- Include summary, priority, reply need, calendar candidates, attachment status
- Open report in browser after generation

### Task 6: Interactive review UI

Implement a simple web UI for interactive review.

Requirements:
- Show one email at a time or grouped list
- Allow mark processed / pending / skipped
- Allow draft generation and editing
- Allow calendar candidate confirmation
- Do not send email without explicit user action

### Task 7: Gmail reply sender

Implement confirmed reply sending.

Requirements:
- Send reply in original Gmail thread
- Require explicit UI action
- Save sent reply log
- Do not auto-send

### Task 8: Google Calendar integration

Implement calendar event creation.

Requirements:
- Create events only after explicit user confirmation
- Use extracted meeting candidates
- Allow user editing of title, time, location, attendees before registration
- Save calendar event ID

### Task 9: Attachment permission flow

Implement attachment permission flow.

Requirements:
- Show attachment metadata
- Ask user before downloading
- Download only approved attachments
- Do not send attachment content to LLM unless explicitly approved

## 9. 未決事項

以下は実装前または Phase 1 完了後に決める．

- UI を Streamlit にするか FastAPI + HTML にするか
- LLM API の種類
- HTML レポート保存先のデフォルト
- 初期版ではメール本文全文をローカル SQLite DB に保存する．ただし `MATOMAIL_STORE_EMAIL_BODY=false` で無効化可能にする
- SQLite DB が大きくなりすぎた場合は，日時付きバックアップフォルダへ退避して新しい DB から再開する
- 処理済み状態を Gmail ラベルとしても反映するか
- 会議登録時のデフォルト所要時間
- 返信草案の文体プリセット
- 日本語・英語メールの自動判定方法

## 10. 推奨初期設定

```env
MATOMAIL_LOOKBACK_DAYS=7
MATOMAIL_REPORT_DIR=./reports
MATOMAIL_DB_PATH=./data/matomail.sqlite3
MATOMAIL_RULES_DB_PATH=./data/matomail_rules.sqlite3
MATOMAIL_DB_BACKUP_DIR=./data/backups
MATOMAIL_DB_MAX_SIZE_MB=512
MATOMAIL_STORE_EMAIL_BODY=true
MATOMAIL_TIMEZONE=Asia/Tokyo
MATOMAIL_MAX_EMAILS_PER_RUN=30
MATOMAIL_AUTO_OPEN_REPORT=true
MATOMAIL_DOWNLOAD_ATTACHMENTS=false
MATOMAIL_SEND_EMAIL_WITHOUT_CONFIRMATION=false
MATOMAIL_CREATE_CALENDAR_WITHOUT_CONFIRMATION=false
MATOMAIL_LLM_MODEL=gpt-5.4-mini
```
