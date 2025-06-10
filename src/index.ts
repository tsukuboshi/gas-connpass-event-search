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

import {
  API_CALL_DELAY,
  CONNPASS_API_BASE_URL,
  EVENT_SHEET_COLUMNS,
  MAX_EVENTS_PER_MESSAGE,
  SPREADSHEET_COLUMNS,
  TIME_FILTERING,
} from './env';

// --- 型定義 ---
interface ConnpassEvent {
  event_id: number;
  title: string;
  catch: string;
  description: string;
  url: string;
  started_at: string;
  ended_at: string;
  limit: number;
  accepted: number;
  waiting: number;
  updated_at: string;
  owner_id: number;
  owner_nickname: string;
  owner_display_name: string;
  place: string;
  address: string;
  lat: string;
  lon: string;
  series: {
    id: number;
    title: string;
    url: string;
  };
}

interface ConnpassApiResponse {
  results_returned: number;
  results_available: number;
  results_start: number;
  events: ConnpassEvent[];
}

interface SpreadsheetConfig {
  keywords: string[];
  connpassApiKey: string;
  lineChannelAccessToken: string;
}

// --- スプレッドシートから設定情報を取得 ---
function getConfigFromSheet(): SpreadsheetConfig {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // A2: キーワード
  const keywordRaw = sheet.getRange(2, SPREADSHEET_COLUMNS.KEYWORDS).getValue();

  // B2: Connpass APIキー
  const connpassApiKeyRaw = sheet
    .getRange(2, SPREADSHEET_COLUMNS.CONNPASS_API_KEY)
    .getValue();

  // C2: LINE Channel Access Token
  const lineChannelAccessTokenRaw = sheet
    .getRange(2, SPREADSHEET_COLUMNS.LINE_CHANNEL_ACCESS_TOKEN)
    .getValue();

  // 値のクリーニング
  const keyword = keywordRaw ? keywordRaw.toString().trim() : '';
  const connpassApiKey = connpassApiKeyRaw
    ? String(connpassApiKeyRaw)
        .trim()
        .replace(/[\r\n]/g, '')
    : '';
  const lineChannelAccessToken = lineChannelAccessTokenRaw
    ? String(lineChannelAccessTokenRaw)
        .trim()
        .replace(/[\r\n]/g, '')
    : '';

  console.log('Cleaned keyword:', keyword);
  console.log('Cleaned Connpass API Key length:', connpassApiKey.length);
  console.log(
    'Cleaned LINE Channel Access Token length:',
    lineChannelAccessToken.length
  );

  // 必要な値が設定されているかチェック
  if (!keyword || keyword === '') {
    throw new Error('検索キーワードが設定されていません（A2セル）');
  }

  if (!connpassApiKey || connpassApiKey === '') {
    throw new Error('Connpass APIキーが設定されていません（B2セル）');
  }

  if (!lineChannelAccessToken || lineChannelAccessToken === '') {
    throw new Error('LINE Channel Access Tokenが設定されていません（C2セル）');
  }

  return {
    keywords: [keyword],
    connpassApiKey,
    lineChannelAccessToken,
  };
}

// --- Connpass APIでイベント検索 ---
function searchConnpassEvents(
  keyword: string,
  apiKey: string
): ConnpassEvent[] {
  try {
    const encodedKeyword = encodeURIComponent(keyword);
    const url = `${CONNPASS_API_BASE_URL}?keyword=${encodedKeyword}`;

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'get',
      headers: {
        'X-API-Key': apiKey,
      },
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `Connpass API エラー: ${response.getResponseCode()} - ${response.getContentText()}`
      );
    }

    const json: ConnpassApiResponse = JSON.parse(response.getContentText());
    const events = json.events || [];

    // デバッグ: 最初の数件のイベントIDの形式を確認
    if (events.length > 0) {
      events.slice(0, 3).forEach((event, index) => {
        console.log(
          `イベント${index + 1}: ID=${event.event_id} (型: ${typeof event.event_id}), タイトル="${event.title}"`
        );
      });
    }

    return events;
  } catch (error) {
    console.error(`Connpass API検索エラー (キーワード: ${keyword}):`, error);
    throw error;
  }
}

// --- LINE Messaging APIでメッセージ送信 ---
function sendLineMessage(message: string, lineToken: string): void {
  try {
    // デバッグ: トークンの確認
    console.log(
      'LINE Channel Access Token length:',
      lineToken ? lineToken.length : 'undefined'
    );
    console.log(
      'LINE Channel Access Token starts with:',
      lineToken ? lineToken.substring(0, 10) + '...' : 'undefined'
    );

    if (!lineToken || lineToken.trim() === '') {
      throw new Error('LINE Channel Access Tokenが設定されていないか、空です');
    }

    // 送信内容
    const payload = {
      messages: [
        {
          type: 'text',
          text: message,
        },
      ],
    };

    // 送信オプション
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + lineToken,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
    };

    console.log('Sending request to LINE API...');

    // 送信
    const response = UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/message/broadcast',
      options
    );

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `LINE API エラー: ${response.getResponseCode()} - ${response.getContentText()}`
      );
    }

    console.log('LINEメッセージ送信成功');
  } catch (error) {
    console.error('LINE メッセージ送信エラー:', error);
    throw error;
  }
}

// --- イベント情報をフォーマット ---
function formatEventMessage(event: ConnpassEvent): string {
  const startDate = new Date(event.started_at);
  const formattedDate = Utilities.formatDate(
    startDate,
    TIME_FILTERING.TIMEZONE,
    'yyyy/MM/dd HH:mm'
  );

  let message = `📅 ${event.title}\n`;
  message += `🕐 ${formattedDate}\n`;

  if (event.place) {
    message += `📍 ${event.place}\n`;
  }

  if (event.accepted !== undefined && event.limit !== undefined) {
    message += `👥 参加者: ${event.accepted}/${event.limit}人\n`;
  }

  message += `🔗 ${event.url}`;

  return message;
}

// --- 過去1時間以内に更新されたイベントをフィルタリング ---
function filterRecentlyUpdatedEvents(events: ConnpassEvent[]): ConnpassEvent[] {
  const now = new Date();
  const oneHourAgo = new Date(
    now.getTime() - TIME_FILTERING.RECENT_HOURS * 60 * 60 * 1000
  );

  console.log(
    `フィルタリング基準時間: ${Utilities.formatDate(oneHourAgo, TIME_FILTERING.TIMEZONE, 'yyyy/MM/dd HH:mm:ss')} 〜 ${Utilities.formatDate(now, TIME_FILTERING.TIMEZONE, 'yyyy/MM/dd HH:mm:ss')}`
  );

  const filteredEvents = events.filter(event => {
    const updatedAt = new Date(event.updated_at);
    const isRecent = updatedAt >= oneHourAgo && updatedAt <= now;

    if (isRecent) {
      console.log(
        `✓ 対象イベント: "${event.title}" (更新: ${Utilities.formatDate(updatedAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')})`
      );
    }

    return isRecent;
  });

  return filteredEvents;
}

// --- 複数イベントをまとめてフォーマット ---
function formatMultipleEvents(
  events: ConnpassEvent[],
  keyword: string
): string {
  if (events.length === 0) {
    return `「${keyword}」に関する過去1時間以内に更新されたイベントは見つかりませんでした。`;
  }

  let message = `🔍 「${keyword}」の検索結果 (過去1時間以内に更新: ${events.length}件)\n\n`;

  events.slice(0, MAX_EVENTS_PER_MESSAGE).forEach((event, index) => {
    message += `${index + 1}. ${formatEventMessage(event)}`;
    if (index < Math.min(events.length - 1, MAX_EVENTS_PER_MESSAGE - 1)) {
      message += '\n\n';
    }
  });

  if (events.length > MAX_EVENTS_PER_MESSAGE) {
    message += `\n\n他 ${events.length - MAX_EVENTS_PER_MESSAGE}件のイベントがあります。`;
  }

  return message;
}

// --- 年月シート管理 ---
function getCurrentYearMonthSheetName(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  return `${year}${month}`;
}

// 前の月のシート名を取得
function getPreviousYearMonthSheetName(): string {
  const now = new Date();
  const previousMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const year = previousMonth.getFullYear();
  const month = String(previousMonth.getMonth() + 1).padStart(2, '0');
  return `${year}${month}`;
}

// 前の月のシートから次の月に開催予定のイベントをコピー
function copyEventsFromPreviousMonth(
  newSheet: GoogleAppsScript.Spreadsheet.Sheet,
  currentSheetName: string
): number {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const previousSheetName = getPreviousYearMonthSheetName();
    const previousSheet = spreadsheet.getSheetByName(previousSheetName);

    if (!previousSheet) {
      console.log(`前の月のシート "${previousSheetName}" が存在しません`);
      return 0;
    }

    const lastRow = previousSheet.getLastRow();
    if (lastRow <= 1) {
      console.log(`前の月のシート "${previousSheetName}" にデータがありません`);
      return 0;
    }

    // 前の月のシートからデータを取得（2行目以降）
    const dataRange = previousSheet.getRange(2, 1, lastRow - 1, 5);
    const values = dataRange.getValues();

    // 現在の年月を取得（YYYY-MM形式）
    const currentYearMonth =
      currentSheetName.substring(0, 4) + '-' + currentSheetName.substring(4, 6);

    let copiedCount = 0;

    console.log(
      `前の月のシート "${previousSheetName}" から ${values.length}件のイベントをチェック中...`
    );
    console.log(`対象年月: ${currentYearMonth}`);

    values.forEach((row, index) => {
      const title = row[EVENT_SHEET_COLUMNS.TITLE - 1];
      const startDateStr = row[EVENT_SHEET_COLUMNS.START_DATE - 1];
      const url = row[EVENT_SHEET_COLUMNS.URL - 1];
      const keyword = row[EVENT_SHEET_COLUMNS.KEYWORD - 1];

      // 開催日時が文字列の場合、Dateオブジェクトに変換
      let startDate: Date;
      if (startDateStr instanceof Date) {
        startDate = startDateStr;
      } else if (
        typeof startDateStr === 'string' &&
        startDateStr.trim() !== ''
      ) {
        // "YYYY/MM/DD HH:mm" 形式の文字列をDateオブジェクトに変換
        startDate = new Date(startDateStr);
      } else {
        console.log(
          `行 ${index + 2}: 開催日時が無効なため、スキップします - ${startDateStr}`
        );
        return;
      }

      // 開催日時が現在の年月に該当するかチェック
      if (isNaN(startDate.getTime())) {
        console.log(
          `行 ${index + 2}: 開催日時の解析に失敗したため、スキップします - ${startDateStr}`
        );
        return;
      }

      const eventYearMonth = Utilities.formatDate(
        startDate,
        TIME_FILTERING.TIMEZONE,
        'yyyy-MM'
      );

      if (eventYearMonth === currentYearMonth) {
        console.log(
          `✓ コピー対象: "${title}" (開催: ${Utilities.formatDate(startDate, TIME_FILTERING.TIMEZONE, 'yyyy/MM/dd HH:mm')})`
        );

        // 新しいシートに追加
        const newRowIndex = newSheet.getLastRow() + 1;

        newSheet
          .getRange(newRowIndex, EVENT_SHEET_COLUMNS.TITLE)
          .setValue(title);
        newSheet
          .getRange(newRowIndex, EVENT_SHEET_COLUMNS.START_DATE)
          .setValue(
            Utilities.formatDate(
              startDate,
              TIME_FILTERING.TIMEZONE,
              'yyyy/MM/dd HH:mm'
            )
          );
        newSheet.getRange(newRowIndex, EVENT_SHEET_COLUMNS.URL).setValue(url);
        newSheet
          .getRange(newRowIndex, EVENT_SHEET_COLUMNS.NOTIFIED_DATE)
          .setValue(
            Utilities.formatDate(
              new Date(),
              TIME_FILTERING.TIMEZONE,
              'yyyy/MM/dd HH:mm:ss'
            ) + ' (前月から引継)'
          );
        newSheet
          .getRange(newRowIndex, EVENT_SHEET_COLUMNS.KEYWORD)
          .setValue(keyword);

        // 枠線を設定
        const rowRange = newSheet.getRange(newRowIndex, 1, 1, 5);
        rowRange.setBorder(true, true, true, true, true, true);

        copiedCount++;
      } else {
        console.log(
          `  スキップ: "${title}" (開催: ${eventYearMonth}) - 対象外の年月`
        );
      }
    });

    if (copiedCount > 0) {
      console.log(
        `前の月のシートから ${copiedCount}件のイベントをコピーしました`
      );
    } else {
      console.log('前の月のシートから該当するイベントはありませんでした');
    }

    return copiedCount;
  } catch (error) {
    console.error('前の月のイベントコピー中にエラーが発生しました:', error);
    return 0;
  }
}

function createOrGetYearMonthSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getCurrentYearMonthSheetName();

  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // シートが存在しない場合は作成
    sheet = spreadsheet.insertSheet(sheetName);

    // ヘッダー行を設定
    sheet.getRange(1, EVENT_SHEET_COLUMNS.TITLE).setValue('タイトル');
    sheet.getRange(1, EVENT_SHEET_COLUMNS.START_DATE).setValue('開催日時');
    sheet.getRange(1, EVENT_SHEET_COLUMNS.URL).setValue('URL');
    sheet.getRange(1, EVENT_SHEET_COLUMNS.NOTIFIED_DATE).setValue('通知日時');
    sheet.getRange(1, EVENT_SHEET_COLUMNS.KEYWORD).setValue('検索キーワード');

    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');
    headerRange.setBorder(true, true, true, true, true, true);

    // 列幅の自動調整
    sheet.autoResizeColumns(1, 5);

    console.log(`年月シート "${sheetName}" を作成しました`);

    // 前の月のシートが存在する場合のみ、該当するイベントをコピー
    const previousSheetName = getPreviousYearMonthSheetName();
    const previousSheet = spreadsheet.getSheetByName(previousSheetName);

    if (previousSheet) {
      console.log(
        `前の月のシート "${previousSheetName}" が存在するため、該当イベントをコピーします`
      );
      const copiedCount = copyEventsFromPreviousMonth(sheet, sheetName);
      if (copiedCount > 0) {
        console.log(
          `新規シート作成時に前の月から ${copiedCount}件のイベントを引き継ぎました`
        );
      }
    } else {
      console.log(
        `前の月のシート "${previousSheetName}" が存在しないため、シートとヘッダーのみを作成しました`
      );
    }
  }

  return sheet;
}

function isEventAlreadyNotified(
  eventUrl: string,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): boolean {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false; // ヘッダー行のみの場合
  }

  // C列のURLを検索
  const eventUrls = sheet
    .getRange(2, EVENT_SHEET_COLUMNS.URL, lastRow - 1, 1)
    .getValues();

  console.log(
    `既存イベントURL検索: 対象URL=${eventUrl}, 既存レコード数=${eventUrls.length}`
  );

  for (let i = 0; i < eventUrls.length; i++) {
    const existingUrl = String(eventUrls[i][0]).trim();
    const targetUrl = String(eventUrl).trim();

    if (existingUrl === targetUrl) {
      console.log(`既存イベント発見: "${existingUrl}"`);
      return true;
    }

    // デバッグ用：一致しなかった場合（最初の3件のみ）
    if (i < 3) {
      console.log(
        `URL比較: 既存="${existingUrl}" vs 対象="${targetUrl}" -> 不一致`
      );
    }
  }

  console.log(`新規イベント: URL=${eventUrl}`);
  return false;
}

function addEventToSheet(
  event: ConnpassEvent,
  keyword: string,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): void {
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;

  const now = new Date();
  const startDate = new Date(event.started_at);

  // デバッグ用ログ
  console.log(
    `シートにイベントを追記中: URL=${event.url}, タイトル="${event.title}"`
  );
  console.log(`追記先行番号: ${newRow}`);

  // イベント情報を追記
  sheet.getRange(newRow, EVENT_SHEET_COLUMNS.TITLE).setValue(event.title);
  sheet
    .getRange(newRow, EVENT_SHEET_COLUMNS.START_DATE)
    .setValue(
      Utilities.formatDate(
        startDate,
        TIME_FILTERING.TIMEZONE,
        'yyyy/MM/dd HH:mm'
      )
    );
  sheet.getRange(newRow, EVENT_SHEET_COLUMNS.URL).setValue(event.url);
  sheet
    .getRange(newRow, EVENT_SHEET_COLUMNS.NOTIFIED_DATE)
    .setValue(
      Utilities.formatDate(now, TIME_FILTERING.TIMEZONE, 'yyyy/MM/dd HH:mm:ss')
    );
  sheet.getRange(newRow, EVENT_SHEET_COLUMNS.KEYWORD).setValue(keyword);

  // 枠線を設定
  const rowRange = sheet.getRange(newRow, 1, 1, 5);
  rowRange.setBorder(true, true, true, true, true, true);

  console.log(
    `イベント "${event.title}" (URL: ${event.url}) をシートの${newRow}行目に追記しました`
  );
}

// --- メインの実行関数 ---
function main(): void {
  try {
    console.log('Connpassイベント検索を開始します...');

    // スプレッドシートから設定を取得
    const config = getConfigFromSheet();
    console.log(`取得したキーワード数: ${config.keywords.length}`);

    // 年月シートを取得または作成
    const eventSheet = createOrGetYearMonthSheet();
    console.log(`年月シート "${getCurrentYearMonthSheetName()}" を使用します`);

    let totalNewEvents = 0;
    let totalAlreadyNotified = 0;

    // 各キーワードで検索
    config.keywords.forEach((keyword, index) => {
      try {
        console.log(
          `${index + 1}/${config.keywords.length}: "${keyword}" を検索中...`
        );

        const events = searchConnpassEvents(keyword, config.connpassApiKey);
        console.log(
          `"${keyword}" で ${events.length}件のイベントが見つかりました`
        );

        // 過去1時間以内に更新されたイベントのみフィルタリング
        const recentlyUpdatedEvents = filterRecentlyUpdatedEvents(events);
        console.log(
          `"${keyword}" で ${recentlyUpdatedEvents.length}件の過去1時間以内に更新されたイベントが見つかりました`
        );

        // 新規イベントと既通知イベントを分類
        const newEvents: ConnpassEvent[] = [];
        const alreadyNotifiedEvents: ConnpassEvent[] = [];

        recentlyUpdatedEvents.forEach(event => {
          if (isEventAlreadyNotified(event.url, eventSheet)) {
            alreadyNotifiedEvents.push(event);
          } else {
            newEvents.push(event);
          }
        });

        // 既に通知済みのイベントがある場合
        if (alreadyNotifiedEvents.length > 0) {
          console.log(
            `"${keyword}" で ${alreadyNotifiedEvents.length}件の既通知済みイベントが見つかりました`
          );

          // 既通知済みイベントの詳細をログに出力（LINE送信はしない）
          alreadyNotifiedEvents
            .slice(0, MAX_EVENTS_PER_MESSAGE)
            .forEach((event, index) => {
              console.log(
                `  既通知済み ${index + 1}: "${event.title}" - ${event.url}`
              );
            });

          if (alreadyNotifiedEvents.length > MAX_EVENTS_PER_MESSAGE) {
            console.log(
              `  他 ${alreadyNotifiedEvents.length - MAX_EVENTS_PER_MESSAGE}件の既通知済みイベントがあります`
            );
          }

          console.log(
            '既通知済みイベントはLINE送信せず、カウントのみ行いました'
          );
          totalAlreadyNotified += alreadyNotifiedEvents.length;
        }

        // 新規イベントがある場合
        if (newEvents.length > 0) {
          console.log(
            `"${keyword}" で ${newEvents.length}件の新規イベントが見つかりました`
          );

          // LINEでメッセージ送信
          const message = formatMultipleEvents(newEvents, keyword);
          sendLineMessage(message, config.lineChannelAccessToken);

          // シートにイベントを追記
          newEvents.forEach(event => {
            addEventToSheet(event, keyword, eventSheet);
          });

          totalNewEvents += newEvents.length;
        } else if (alreadyNotifiedEvents.length === 0) {
          console.log(`"${keyword}" で新規イベントはありませんでした`);
        }

        // API制限を考慮して少し待機
        Utilities.sleep(API_CALL_DELAY);
      } catch (error) {
        console.error(
          `キーワード "${keyword}" の処理中にエラーが発生しました:`,
          error
        );
        // 個別のキーワードでエラーが発生しても他のキーワードの処理は続行
      }
    });

    // 結果のメッセージを作成
    let resultMessage = '';
    if (totalNewEvents === 0 && totalAlreadyNotified === 0) {
      resultMessage =
        '過去1時間以内に更新された新しいイベントは見つかりませんでした。LINEメッセージは送信しません。';
    } else {
      resultMessage = `処理完了: 新規イベント ${totalNewEvents}件を通知、既通知済みイベント ${totalAlreadyNotified}件を確認しました`;
    }

    console.log(resultMessage);

    // スプレッドシートのUIにも結果を表示
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('検索結果', resultMessage, ui.ButtonSet.OK);
    } catch (uiError) {
      console.log(
        'UI表示エラー（スクリプトエディタから実行された可能性があります）:',
        uiError
      );
    }
  } catch (error) {
    console.error('メイン処理でエラーが発生しました:', error);

    // エラーが発生した場合はUIにのみ表示（LINEには送信しない）
    try {
      const ui = SpreadsheetApp.getUi();
      const errorMessage = `❌ Connpassイベント検索でエラーが発生しました:\n${error instanceof Error ? error.message : String(error)}`;
      ui.alert('エラー', errorMessage, ui.ButtonSet.OK);
    } catch (uiError) {
      console.log(
        'UI表示エラー（スクリプトエディタから実行された可能性があります）:',
        uiError
      );
    }

    throw error;
  }
}

// --- LINE Token形式チェック関数 ---
function validateLineToken(token: string): {
  isValid: boolean;
  message: string;
} {
  if (!token || token.trim() === '') {
    return { isValid: false, message: 'トークンが空です' };
  }

  // LINE Channel Access Tokenの一般的な形式をチェック
  // 通常は英数字と記号で構成され、長さは40文字以上
  if (token.length < 40) {
    return {
      isValid: false,
      message: `トークンが短すぎます (${token.length}文字)`,
    };
  }

  // 基本的な文字チェック（英数字、+、/、=のみ許可）
  const validChars = /^[A-Za-z0-9+/=]+$/;
  if (!validChars.test(token)) {
    return { isValid: false, message: '無効な文字が含まれています' };
  }

  return { isValid: true, message: 'トークン形式は正常です' };
}

// --- 統合テスト関数 ---
function testEvents(
  maxEvents: number = 5,
  sendTestMessage: boolean = true
): void {
  try {
    console.log(
      `統合テストを開始します（処理件数: ${maxEvents}件、テストメッセージ: ${sendTestMessage ? 'あり' : 'なし'}）...`
    );

    // 1. 設定情報の検証
    const config = getConfigFromSheet();
    console.log('設定情報の取得に成功しました');
    console.log(`キーワード数: ${config.keywords.length}`);
    console.log(`キーワード: ${config.keywords.join(', ')}`);

    // デバッグ: 設定値の確認
    console.log(
      `Connpass API Key length:`,
      config.connpassApiKey ? config.connpassApiKey.length : 'undefined'
    );
    console.log(
      `LINE Channel Access Token length:`,
      config.lineChannelAccessToken
        ? config.lineChannelAccessToken.length
        : 'undefined'
    );

    // LINE Channel Access Tokenの詳細チェック
    const tokenValidation = validateLineToken(config.lineChannelAccessToken);
    console.log(
      `LINE Channel Access Token validation:`,
      tokenValidation.message
    );

    if (!tokenValidation.isValid) {
      throw new Error(
        `LINE Channel Access Token検証エラー: ${tokenValidation.message}`
      );
    }

    // 2. 年月シートの準備
    const eventSheet = createOrGetYearMonthSheet();
    console.log(`年月シート "${getCurrentYearMonthSheetName()}" を使用します`);

    // 3. テストメッセージ送信（オプション）
    if (sendTestMessage) {
      const testMessage =
        '🧪 GAS Connpassイベント検索アプリのテストメッセージです。';
      sendLineMessage(testMessage, config.lineChannelAccessToken);
      console.log('テストメッセージの送信に成功しました');
    }

    // 4. 検索とイベント処理
    const keyword = config.keywords[0];
    console.log(`テスト検索キーワード: "${keyword}"`);

    const events = searchConnpassEvents(keyword, config.connpassApiKey);
    console.log(`"${keyword}" で ${events.length}件のイベントが見つかりました`);

    if (events.length > 0) {
      // 指定件数分のイベントを処理
      const testEvents = events.slice(0, Math.min(events.length, maxEvents));
      console.log(`テスト対象イベント: ${testEvents.length}件`);

      // 新規イベントと既通知イベントを分類
      const newEvents: ConnpassEvent[] = [];
      const alreadyNotifiedEvents: ConnpassEvent[] = [];

      testEvents.forEach((event, index) => {
        console.log(
          `イベント ${index + 1}/${testEvents.length}: ID=${event.event_id}, タイトル="${event.title}"`
        );
        console.log(`更新日時: ${event.updated_at}`);

        if (isEventAlreadyNotified(event.url, eventSheet)) {
          console.log(`  → 既通知済み`);
          alreadyNotifiedEvents.push(event);
        } else {
          console.log(`  → 新規イベント`);
          newEvents.push(event);
        }
      });

      // 新規イベントの処理
      if (newEvents.length > 0) {
        console.log(`${newEvents.length}件の新規イベントを処理します`);

        // LINE送信用メッセージを作成
        let searchResultMessage = `🔍 「${keyword}」のテスト検索結果 (新規: ${newEvents.length}件)\n\n`;

        newEvents.forEach((event, index) => {
          searchResultMessage += `${index + 1}. ${formatEventMessage(event)}`;
          if (index < newEvents.length - 1) {
            searchResultMessage += '\n\n';
          }
        });

        sendLineMessage(searchResultMessage, config.lineChannelAccessToken);
        console.log('新規イベントの検索結果を送信しました');

        // シートにイベントを追記
        newEvents.forEach((event, index) => {
          console.log(
            `新規イベント ${index + 1}/${newEvents.length} をシートに追記中...`
          );
          addEventToSheet(event, keyword, eventSheet);
        });
      }

      // 既通知済みイベントの処理
      if (alreadyNotifiedEvents.length > 0) {
        console.log(
          `${alreadyNotifiedEvents.length}件の既通知済みイベントが見つかりました`
        );

        // 既通知済みイベントの詳細をログに出力（LINE送信はしない）
        alreadyNotifiedEvents.forEach((event, index) => {
          console.log(
            `  既通知済み ${index + 1}: "${event.title}" - ${event.url}`
          );
        });

        console.log('既通知済みイベントはLINE送信せず、ログ出力のみ行いました');
      }

      // 結果サマリー
      const summaryMessage = `テスト完了: 新規追加 ${newEvents.length}件、既通知済み ${alreadyNotifiedEvents.length}件`;
      console.log(summaryMessage);

      // UIアラート表示
      try {
        const ui = SpreadsheetApp.getUi();
        const resultMessage =
          `統合テスト完了\n` +
          `検索結果: ${events.length}件\n` +
          `処理対象: ${testEvents.length}件\n` +
          `新規追加: ${newEvents.length}件\n` +
          `既通知済み: ${alreadyNotifiedEvents.length}件\n` +
          `年月シート: "${getCurrentYearMonthSheetName()}"`;
        ui.alert('統合テスト結果', resultMessage, ui.ButtonSet.OK);
      } catch (uiError) {
        console.log(
          'UI表示エラー（スクリプトエディタから実行された可能性があります）:',
          uiError
        );
      }
    } else {
      const noResultMessage = `🔍 「${keyword}」の検索結果: イベントが見つかりませんでした。`;
      sendLineMessage(noResultMessage, config.lineChannelAccessToken);
      console.log('検索結果なしメッセージの送信に成功しました');

      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          '統合テスト結果',
          `検索結果が見つかりませんでした（キーワード: ${keyword}）`,
          ui.ButtonSet.OK
        );
      } catch (uiError) {
        console.log('UI表示エラー:', uiError);
      }
    }
  } catch (error) {
    console.error('統合テストでエラーが発生しました:', error);
    throw error;
  }
}

// --- スプレッドシート初期化関数 ---
function initializeSpreadsheet(): void {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // シート名を「プロパティ」に変更（Sheet1の場合のみ）
    if (sheet.getName() === 'Sheet1' || sheet.getName().startsWith('シート')) {
      sheet.setName('プロパティ');
      console.log('シート名を「プロパティ」に変更しました');
    }

    // シートをクリア
    sheet.clear();

    // 1行目: ヘッダー
    sheet.getRange(1, 1).setValue('検索キーワード');
    sheet.getRange(1, 2).setValue('Connpass APIキー');
    sheet.getRange(1, 3).setValue('LINE Channel Access Token');

    // 2行目: サンプルキーワードを設定
    sheet.getRange(2, 1).setValue('typescript');
    // B2とC2は空のままにして、ユーザーが直接入力できるようにする

    // ヘッダー行の書式設定（1行目）
    const headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');

    // セルにコメントを追加（背景色は設定しない）
    sheet.getRange(2, 1).setNote('検索キーワードを入力してください');
    sheet.getRange(2, 2).setNote('Connpass APIキーを入力してください');
    sheet.getRange(2, 3).setNote('LINE Channel Access Tokenを入力してください');

    // 列幅の自動調整
    sheet.autoResizeColumns(1, 3);

    // 枠線の設定（2行分）
    const dataRange = sheet.getRange(1, 1, 2, 3);
    dataRange.setBorder(true, true, true, true, true, true);

    // 説明の追加
    const instructionStartRow = 4;
    sheet.getRange(instructionStartRow, 1).setValue('設定方法:');
    sheet.getRange(instructionStartRow + 1, 1).setValue('• A2: 検索キーワード');
    sheet
      .getRange(instructionStartRow + 2, 1)
      .setValue('• B2: Connpass APIキー');
    sheet
      .getRange(instructionStartRow + 3, 1)
      .setValue('• C2: LINE Channel Access Token');

    const instructionRange = sheet.getRange(instructionStartRow, 1, 4, 1);
    instructionRange.setFontStyle('italic');
    instructionRange.setFontColor('#666666');

    console.log('スプレッドシートの初期化が完了しました');
    console.log(
      'B2セルにConnpass APIキー、C2セルにLINE Channel Access Tokenを入力してください。'
    );
  } catch (error) {
    console.error('スプレッドシート初期化エラー:', error);
    throw error;
  }
}

// --- スプレッドシート自動初期化（onOpen時に実行） ---
function onOpen(): void {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // シート名を「プロパティ」に変更（Sheet1の場合のみ）
    if (sheet.getName() === 'Sheet1' || sheet.getName().startsWith('シート')) {
      sheet.setName('プロパティ');
      console.log('シート名を「プロパティ」に変更しました');
    }

    // シートが空の場合のみ初期化
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === '') {
      initializeSpreadsheet();
      console.log('スプレッドシートの初期化が完了しました');
    }

    // カスタムメニューの追加
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Connpassイベント検索')
      .addItem('設定を初期化', 'initializeSpreadsheet')
      .addItem('接続・動作確認テスト', 'testEvents')
      .addItem('過去1時間以内イベント検索', 'main')
      .addToUi();
  } catch (error) {
    console.error('onOpen実行エラー:', error);
  }
}
