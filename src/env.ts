/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * Copyright 2023 tsukuboshi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// Connpass API設定
export const CONNPASS_API_BASE_URL = 'https://connpass.com/api/v2/events/';

// LINE Messaging API設定
export const LINE_API_BASE_URL = 'https://api.line.me/v2/bot/message/broadcast';

// 一度に表示するイベントの最大数
export const MAX_EVENTS_PER_MESSAGE = 5;

// 保存する既知のイベントIDの最大数
export const MAX_KNOWN_EVENT_IDS = 1000;

// API呼び出し間の待機時間（ミリ秒）
export const API_CALL_DELAY = 1000;

// スプレッドシートの列定義
export const SPREADSHEET_COLUMNS = {
  CONNPASS_API_KEY: 1, // A列: Connpass APIキー
  LINE_CHANNEL_ACCESS_TOKEN: 2, // B列: LINE Channel Access Token
  KEYWORDS: 3, // C列: キーワード
} as const;

// 年月シートの列定義
export const EVENT_SHEET_COLUMNS = {
  TITLE: 1, // A列: タイトル
  START_DATE: 2, // B列: 開催日時
  URL: 3, // C列: URL
  NOTIFIED_DATE: 4, // D列: 通知日時
  KEYWORD: 5, // E列: 検索キーワード
} as const;

// スプレッドシートの行定義
export const SPREADSHEET_ROWS = {
  CONFIG_ROW: 1, // 1行目: 設定情報
  KEYWORDS_START: 2, // 2行目以降: キーワード
} as const;

// 年月シート設定
export const YEAR_MONTH_SHEET = {
  NAME_FORMAT: 'YYYYMM', // シート名の形式
  HEADER_ROW: 1, // ヘッダー行
  DATA_START_ROW: 2, // データ開始行
} as const;

// 時間フィルタリング設定
export const TIME_FILTERING = {
  RECENT_HOURS: 1, // 過去何時間以内のイベントを対象とするか
  TIMEZONE: 'Asia/Tokyo', // タイムゾーン
} as const;
