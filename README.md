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

このGoogle Apps Script（GAS）アプリケーションは、Connpass APIを使用してイベントを検索し、結果をLINE Messaging APIで通知するサンプルアプリケーションです。

## 機能

- スプレッドシートに記載されたキーワードでConnpassイベントを検索
- Connpass APIキーを使用した認証付きAPI呼び出し
- 検索結果をLINE Messaging APIで自動通知
- 新規イベントのみを通知（重複通知を防止）
- エラーハンドリングとログ出力

## セットアップ

### 1. スプレッドシートの準備

Google スプレッドシートを作成し、GASプロジェクトと関連付けてください。スプレッドシートを開くと、自動的に以下の形式でテンプレートが設定されます：

| A列（キーワード） | B列（Connpass APIキー） | C列（LINE Token） |
|------------------|------------------------|-------------------|
| python           | YOUR_CONNPASS_API_KEY  | YOUR_LINE_TOKEN   |
| javascript       |                        |                   |
| react            |                        |                   |
| vue              |                        |                   |
| typescript       |                        |                   |

#### 自動初期化機能

- スプレッドシートを初回開いた時に、自動的にテンプレートが設定されます
- ヘッダー行とサンプルキーワードが自動入力されます
- 「Connpassイベント検索」メニューが追加され、以下の機能が利用できます：
  - **設定を初期化**: テンプレートを再設定
  - **接続テスト**: 設定の確認とテスト通知
  - **イベント検索実行**: メイン機能の実行

#### 設定方法

- **A列**: 検索キーワード（2行目以降に複数設定可能）
- **B1セル**: Connpass APIキー（`YOUR_CONNPASS_API_KEY`を実際のキーに置き換え）
- **C1セル**: LINE Messaging API のチャネルアクセストークン（`YOUR_LINE_TOKEN`を実際のトークンに置き換え）

#### 手動設定（必要に応じて）

自動初期化が動作しない場合は、GASエディタで`initializeSpreadsheet()`関数を実行するか、上記の表の形式で手動入力してください。

### 2. Connpass APIキーの取得

1. [Connpass](https://connpass.com/) にログイン
2. APIキーを取得（詳細は Connpass の API ドキュメントを参照）
3. 取得したAPIキーをスプレッドシートのB1セルに入力

### 3. LINE Messaging APIの設定

1. [LINE Developers Console](https://developers.line.biz/) でチャネルを作成
2. Messaging API のチャネルアクセストークンを取得
3. 取得したトークンをスプレッドシートのC1セルに入力

### 4. GASプロジェクトの設定

1. Google Apps Script プロジェクトを作成
2. このリポジトリのコードをデプロイ
3. スプレッドシートをGASプロジェクトに関連付け

## 使用方法

### 手動実行

```javascript
// メイン関数を実行
main();

// 接続テスト
testConnection();
```

### 定期実行（トリガー設定）

1. GASエディタで「トリガー」を設定
2. `main` 関数を定期実行するよう設定
3. 実行頻度は用途に応じて調整（例：1時間ごと、1日1回など）

## API仕様

### Connpass API

```bash
curl -X GET "https://connpass.com/api/v2/events/?keyword=python" \
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

メインの実行関数。スプレッドシートから設定を読み込み、各キーワードでイベント検索を実行し、新規イベントをLINEで通知します。

### `getConfigFromSheet()`

スプレッドシートから設定情報（キーワード、APIキー、LINEトークン）を取得します。

### `searchConnpassEvents(keyword, apiKey)`

指定されたキーワードでConnpass APIを呼び出し、イベント情報を取得します。

### `sendLineMessage(message, lineToken)`

LINE Messaging APIを使用してメッセージを送信します。

### `testConnection()`

設定の確認とLINE通知のテストを行います。

### `initializeSpreadsheet()`

スプレッドシートにテンプレートを設定します。ヘッダー、サンプルキーワード、書式設定を自動で行います。

### `onOpen()`

スプレッドシートを開いた時に自動実行される関数。空のシートの場合は自動初期化を行い、カスタムメニューを追加します。

## エラーハンドリング

- 設定値の不備（APIキー、トークンの未設定など）
- API呼び出しエラー（認証エラー、レート制限など）
- ネットワークエラー
- データ形式エラー

エラーが発生した場合は、可能であればLINEでエラー通知を送信します。

## 制限事項

- Connpass APIのレート制限に注意
- LINE Messaging APIの月間メッセージ数制限
- 一度に表示するイベント数は5件まで
- 保存する既知のイベントIDは最新1000件まで

## ライセンス

Apache License 2.0

## 注意事項

- このアプリケーションはサンプル実装です
- 本番環境で使用する場合は、適切なエラーハンドリングとセキュリティ対策を実装してください
- APIキーやトークンは適切に管理し、外部に漏洩しないよう注意してください
