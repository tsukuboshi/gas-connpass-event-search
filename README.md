<!--
Copyright 2023 Google LLC

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
-->
# Connpass イベント検索 GAS アプリケーション

このGoogle Apps Script（GAS）アプリケーションは、Connpass APIを使用してイベントを検索し、過去1時間以内に更新された新規イベントのみをLINE Messaging APIで通知するアプリケーションです。

## 機能

- スプレッドシートに記載されたキーワードでConnpassイベントを検索
- Connpass APIキーを使用した認証付きAPI呼び出し
- **過去1時間以内に更新されたイベントのみをフィルタリング**
- **新規イベントのみを通知（重複通知を防止）**
- **新規イベントがない場合は通知なし**
- 検索結果をLINE Messaging APIで自動通知
- スプレッドシート自動初期化機能
- カスタムメニューによる操作

## セットアップ

### 1. スプレッドシートの準備

Google スプレッドシートを作成し、GASプロジェクトと関連付けてください。スプレッドシートを開くと、自動的に以下の形式でテンプレートが設定されます：

| 行 | A列（検索キーワード） | B列（Connpass APIキー） | C列（LINE Channel Access Token） |
|---|---------------------|------------------------|----------------------------------|
| **1行目** | 検索キーワード | Connpass APIキー | LINE Channel Access Token |
| **2行目** | typescript | **(空)** | **(空)** |

#### 自動初期化機能

- スプレッドシートを初回開いた時に、自動的にテンプレートが設定されます
- ヘッダー行とサンプルキーワード（typescript）が自動入力されます
- 「Connpassイベント検索」メニューが追加され、以下の機能が利用できます：
  - **設定を初期化**: テンプレートを再設定
  - **検索結果送付テスト**: 設定の確認と全検索結果の送付テスト
  - **検索結果フィルタリングテスト**: 過去1時間以内の新規イベントのみ送付

#### 設定方法

- **A2セル**: 検索キーワード（1つのキーワードのみ）
- **B2セル**: Connpass APIキー
- **C2セル**: LINE Channel Access Token

### 2. Connpass APIキーの取得

1. [Connpass](https://connpass.com/) にログイン
2. APIキーを取得（詳細は Connpass の API ドキュメントを参照）
3. 取得したAPIキーをスプレッドシートのB2セルに入力

### 3. LINE Messaging APIの設定

1. [LINE Developers Console](https://developers.line.biz/) でチャネルを作成
2. Messaging API のチャネルアクセストークンを取得
3. 取得したトークンをスプレッドシートのC2セルに入力

### 4. GASプロジェクトの設定

1. Google Apps Script プロジェクトを作成
2. このリポジトリのコードをデプロイ
3. スプレッドシートをGASプロジェクトに関連付け

## 使用方法

### スプレッドシートメニューから実行

1. **設定を初期化**: スプレッドシートのテンプレートを再設定
2. **検索結果送付テスト**:
   - 設定値の確認
   - テストメッセージの送信
   - 全検索結果の送付（時間フィルタリングなし）
3. **検索結果フィルタリングテスト**:
   - 過去1時間以内に更新された新規イベントのみ送付
   - 新規イベントがない場合は通知なし

### 定期実行（トリガー設定）

1. GASエディタで「トリガー」を設定
2. `main` 関数を**1時間ごと**に実行するよう設定
3. 過去1時間以内に更新された新規イベントのみが通知されます

### 手動実行（GASエディタから）

```javascript
// メイン関数を実行（過去1時間以内の新規イベントのみ）
main();

// 検索結果送付テスト（全検索結果）
testConnection();
```

## 時間フィルタリング機能

### 動作仕様

- **実行時刻から1時間前まで**に更新されたイベントのみを対象
- Connpass APIの`updated_at`フィールドを使用
- 新規イベント（過去に通知していないイベント）のみを通知
- 新規イベントがない場合はLINE通知なし（ログのみ）

### 例

- 14:00に実行 → 13:00〜14:00に更新されたイベントが対象
- 15:00に実行 → 14:00〜15:00に更新されたイベントが対象

## API仕様

### Connpass API

```bash
curl -X GET "https://connpass.com/api/v2/events/?keyword=typescript" \
-H "X-API-Key: YOUR_API_KEY"
```

### LINE Messaging API

```bash
curl -X POST "https://api.line.me/v2/bot/message/broadcast" \
-H "Authorization: Bearer YOUR_CHANNEL_ACCESS_TOKEN" \
-H "Content-Type: application/json" \
-d '{
  "messages": [
    {
      "type": "text",
      "text": "メッセージ内容"
    }
  ]
}'
```

## 主要な関数

### `main()`

メインの実行関数。過去1時間以内に更新された新規イベントのみをLINEで通知します。

### `getConfigFromSheet()`

スプレッドシートから設定情報（A2: キーワード、B2: APIキー、C2: LINEトークン）を取得します。

### `searchConnpassEvents(keyword, apiKey)`

指定されたキーワードでConnpass APIを呼び出し、イベント情報を取得します。

### `filterRecentlyUpdatedEvents(events)`

過去1時間以内に更新されたイベントのみをフィルタリングします。

### `isNewEvent(eventId)`

新規イベント判定を行い、既知のイベントIDを管理します（最新1000件まで保存）。

### `sendLineMessage(message, lineToken)`

LINE Messaging APIを使用してブロードキャストメッセージを送信します。

### `testConnection()`

設定の確認、テストメッセージの送信、全検索結果の送付を行います（時間フィルタリングなし）。

### `initializeSpreadsheet()`

スプレッドシートにテンプレートを設定します。ヘッダー、サンプルキーワード、書式設定を自動で行います。

### `onOpen()`

スプレッドシートを開いた時に自動実行される関数。空のシートの場合は自動初期化を行い、カスタムメニューを追加します。

### `validateLineToken(token)`

LINE Channel Access Tokenの形式をチェックします。

## エラーハンドリング

- 設定値の不備（APIキー、トークンの未設定など）
- API呼び出しエラー（認証エラー、レート制限など）
- ネットワークエラー
- データ形式エラー

**エラーが発生した場合は、LINEには通知せず、ログとUIアラートのみで報告されます。**

## 制限事項

- Connpass APIのレート制限に注意
- LINE Messaging APIの月間メッセージ数制限
- 一度に表示するイベント数は5件まで（`MAX_EVENTS_PER_MESSAGE`）
- 保存する既知のイベントIDは最新1000件まで（`MAX_KNOWN_EVENT_IDS`）
- 検索キーワードは1つのみ（A2セル）
- API呼び出し間隔は1秒（`API_CALL_DELAY`）

## メッセージ形式

### 検索結果送付テスト時

```
🧪 GAS Connpassイベント検索アプリのテストメッセージです。

🔍 「typescript」の検索結果 (全件: 10件)

1. 📅 TypeScript勉強会
🕐 2024/01/15 19:00
📍 東京都渋谷区
👥 参加者: 25/30人
🔗 https://connpass.com/event/12345/

2. 📅 React + TypeScript ハンズオン
🕐 2024/01/20 14:00
📍 オンライン
👥 参加者: 45/50人
🔗 https://connpass.com/event/12346/

他 8件のイベントがあります。
```

### 検索結果フィルタリングテスト時

```
🔍 「typescript」の検索結果 (過去1時間以内に更新: 2件)

1. 📅 TypeScript勉強会
🕐 2024/01/15 19:00
📍 東京都渋谷区
👥 参加者: 25/30人
🔗 https://connpass.com/event/12345/

2. 📅 React + TypeScript ハンズオン
🕐 2024/01/20 14:00
📍 オンライン
👥 参加者: 45/50人
🔗 https://connpass.com/event/12346/
```

## ライセンス

Apache License 2.0
