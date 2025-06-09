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

このGoogle Apps Script（GAS）アプリケーションは、Connpass APIを使用してイベントを検索し、**年月シートで重複管理**を行いながら、過去1時間以内に更新された新規イベントのみをLINE Messaging APIで通知するアプリケーションです。

## 主な機能

- **年月シート自動作成**: YYYYMM形式のシートを自動生成・管理
- **URL基準の重複チェック**: イベントURLで既通知済みイベントを判定
- **時間フィルタリング**: 過去1時間以内に更新されたイベントのみを対象
- **新規イベントのみ通知**: 既通知済みイベントはLINE送信せず、ログのみ出力
- **複数キーワード対応**: スプレッドシートの複数行に対応
- **設定自動初期化**: 初回開時の自動テンプレート設定
- **カスタムメニュー**: 分かりやすい操作メニュー

## 📋 スプレッドシート構成

### プロパティシート（設定用）

スプレッドシートを開くと、最初のシートが自動的に「プロパティ」に名前変更され、以下のテンプレートが設定されます：

| 行 | A列（検索キーワード） | B列（Connpass APIキー） | C列（LINE Channel Access Token） |
|---|---------------------|------------------------|----------------------------------|
| **1行目** | 検索キーワード | Connpass APIキー | LINE Channel Access Token |
| **2行目** | typescript | **(設定値)** | **(設定値)** |
| **3行目** | react | | |
| **4行目** | python | | |

### 年月シート（イベント管理用）

現在の年月（例：202412）を基準に自動作成されるシートで、以下の構成：

| A列 | B列 | C列 | D列 | E列 |
|-----|-----|-----|-----|-----|
| **タイトル** | **開始日時** | **URL** | **通知日時** | **キーワード** |
| TypeScript勉強会 | 2024-12-15 19:00 | https://connpass.com/event/12345/ | 2024-12-01 10:30:15 | typescript |

## 🚀 セットアップ

### 1. 必要なAPI設定

#### Connpass APIキーの取得

1. [Connpass API利用ガイド](https://help.connpass.com/api/#id4)を参照
2. 利用申請フォームから申し込み
3. 取得したAPIキーをプロパティシートのB列に入力

#### LINE Messaging APIの設定

1. [LINE Developers Console](https://developers.line.biz/)でチャネル作成
2. チャネルアクセストークンを取得
3. 取得したトークンをプロパティシートのC列に入力

### 2. GASプロジェクトの設定

1. Google Apps Scriptプロジェクトを作成
2. このリポジトリのコードをデプロイ
3. スプレッドシートをGASプロジェクトに関連付け
4. 初回開時に自動的にテンプレートが設定されます

## 📝 使用方法

### カスタムメニューから実行

スプレッドシートに「Connpassイベント検索」メニューが追加され、以下の機能が利用できます：

| メニュー項目 | 実行内容 | 用途 |
|-------------|----------|------|
| **設定を初期化** | プロパティシートのテンプレート再設定 | 初回セットアップ、設定リセット |
| **接続・動作確認テスト** | 設定検証 + テストメッセージ + 全期間検索（5件） | API接続確認、設定値テスト、基本動作確認 |
| **過去1時間以内イベント検索** | 時間フィルタリング + 新規イベント通知 | 本番運用、定期実行確認 |

### 定期実行の設定

1. GASエディタで「トリガー」を設定
2. `main`関数を**1時間ごと**に実行するよう設定
3. 過去1時間以内に更新された新規イベントのみが自動通知されます

## ⚙️ 重複管理システム

### 年月シートでの管理

- **シート名**: YYYYMM形式（例：202412）
- **自動作成**: 現在の年月シートが存在しない場合に自動生成
- **重複判定**: イベントURLを基準に既通知済みを判定
- **データ保持**: 各月のイベント履歴を永続保存

### URL基準の重複チェック

```javascript
// 重複チェックの仕組み
const isAlreadyNotified = (eventUrl: string): boolean => {
  // 年月シートのC列（URL列）を検索
  // マッチするURLがあれば既通知済み
}
```

## 🔍 時間フィルタリング機能

### 対象期間

- **実行時刻から1時間前まで**に更新されたイベントのみ
- Connpass APIの`updated_at`フィールドを使用

### フィルタリング例

```
14:00実行 → 13:00〜14:00に更新されたイベントが対象
15:00実行 → 14:00〜15:00に更新されたイベントが対象
```

### 通知条件

- ✅ **新規イベント**: 年月シートに記録されていないイベント
- ❌ **既通知済み**: 年月シートに既に記録されているイベント
- ❌ **時間外更新**: 過去1時間以外に更新されたイベント

## 📱 メッセージ形式

### 接続・動作確認テスト時

```
🧪 GAS Connpassイベント検索アプリのテストメッセージです。

🔍 「typescript」の検索結果 (新規: 2件)

1. 📅 TypeScript勉強会
🕐 2024/12/15 19:00
📍 東京都渋谷区
👥 参加者: 25/30人
🔗 https://connpass.com/event/12345/

2. 📅 React + TypeScript ハンズオン
🕐 2024/12/20 14:00
📍 オンライン
👥 参加者: 45/50人
🔗 https://connpass.com/event/12346/
```

### 過去1時間以内イベント検索時

```
🔍 「typescript」の検索結果 (過去1時間以内更新: 1件)

1. 📅 TypeScript勉強会
🕐 2024/12/15 19:00
📍 東京都渋谷区
👥 参加者: 25/30人
🔗 https://connpass.com/event/12345/
```

## 🛠️ 主要な関数

### コア機能

- **`main()`**: メイン実行関数（過去1時間以内の新規イベント通知）
- **`testEvents()`**: 統合テスト関数（設定確認 + 全期間検索）

### 設定・初期化

- **`onOpen()`**: スプレッドシート開時の自動初期化
- **`initializeSpreadsheet()`**: プロパティシートのテンプレート設定
- **`getConfigFromSheet()`**: 設定情報の取得

### 年月シート管理

- **`getCurrentYearMonthSheetName()`**: 現在の年月シート名を取得
- **`createOrGetYearMonthSheet()`**: 年月シートの作成または取得
- **`isEventAlreadyNotified()`**: URL基準の重複チェック
- **`addEventToSheet()`**: 新規イベントのシート追記

### API連携

- **`searchConnpassEvents()`**: Connpass API呼び出し
- **`sendLineMessage()`**: LINE Messaging API送信
- **`filterRecentlyUpdatedEvents()`**: 時間フィルタリング

## ⚠️ 制限事項

### API制限

- **Connpass API**: レート制限に注意
- **LINE Messaging API**: 月間メッセージ数制限
- **API呼び出し間隔**: 1秒（`API_CALL_DELAY`）

### 処理制限

- **同時表示イベント数**: 5件まで（`MAX_EVENTS_PER_MESSAGE`）
- **キーワード数**: 制限なし（プロパティシートの行数分）
- **年月シート**: 月ごとに自動分割

### 通知制限

- **既通知済みイベント**: LINE送信なし（ログのみ）
- **時間外イベント**: 過去1時間以外は対象外
- **エラー時**: LINEには通知せず、ログとUIアラートのみ

## 🔧 設定値（env.ts）

```typescript
// API設定
export const CONNPASS_API_BASE_URL = 'https://connpass.com/api/v2/events/';
export const API_CALL_DELAY = 1000; // 1秒

// メッセージ制限
export const MAX_EVENTS_PER_MESSAGE = 5;

// 時間フィルタリング
export const TIME_FILTERING = {
  ENABLED: true,
  HOURS_BACK: 1
};

// イベントシート列定義
export const EVENT_SHEET_COLUMNS = {
  TITLE: 1,        // A列: タイトル
  START_DATE: 2,   // B列: 開始日時
  URL: 3,          // C列: URL
  NOTIFIED_DATE: 4, // D列: 通知日時
  KEYWORD: 5       // E列: キーワード
};
```

## 📊 デバッグ・ログ

### ログ出力内容

- 設定値の読み込み状況
- API呼び出し結果
- 時間フィルタリング結果
- 重複チェック結果
- 年月シート操作状況
- エラー詳細情報

### UIアラート

- 処理結果の統計情報
- 新規追加件数
- 既通知済み件数
- エラー発生時の詳細

## 🧪 テスト機能

### 接続・動作確認テスト

```javascript
// デフォルト実行（5件、テストメッセージあり）
testEvents();

// カスタム実行（3件、テストメッセージなし）
testEvents(3, false);
```

### 本番実行テスト

```javascript
// 過去1時間以内の新規イベントのみ
main();
```

## 📈 運用シナリオ

### 初回セットアップ

1. **設定を初期化** → プロパティシート準備
2. API設定（ConnpassキーとLINEトークン）入力
3. **接続・動作確認テスト** → 基本動作確認
4. **過去1時間以内イベント検索** → 本番動作確認

### 日常運用

- トリガー設定で`main`関数を1時間ごとに自動実行
- 手動実行時は「過去1時間以内イベント検索」メニューを使用

### トラブルシューティング

- **接続・動作確認テスト**で設定値とAPI接続を確認
- 年月シートで過去の通知履歴を確認
- ログとUIアラートでエラー詳細を確認

## ライセンス

Apache License 2.0
