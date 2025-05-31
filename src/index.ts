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

// --- 定数定義 ---
const CONNPASS_API_BASE_URL = 'https://connpass.com/api/v2/events/';
const MAX_EVENTS_PER_MESSAGE = 5;
const MAX_KNOWN_EVENT_IDS = 1000;
const API_CALL_DELAY = 1000;

const SPREADSHEET_COLUMNS = {
  KEYWORDS: 1, // A列: キーワード
  CONNPASS_API_KEY: 2, // B列: Connpass APIキー
  LINE_CHANNEL_ACCESS_TOKEN: 3, // C列: LINE Channel Access Token
} as const;

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
    return json.events || [];
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
    'Asia/Tokyo',
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
  const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000); // 1時間前

  console.log(
    `フィルタリング基準時間: ${Utilities.formatDate(oneHourAgo, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')} 〜 ${Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}`
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

// --- 新規イベント判定・ID管理 ---
function isNewEvent(eventId: number): boolean {
  const props = PropertiesService.getScriptProperties();
  const known = props.getProperty('knownEventIds');
  const knownIds = known ? known.split(',').map(Number) : [];

  if (knownIds.includes(eventId)) {
    return false;
  }

  knownIds.push(eventId);
  // 保存するIDの数を制限（最新1000件まで）
  if (knownIds.length > MAX_KNOWN_EVENT_IDS) {
    knownIds.splice(0, knownIds.length - MAX_KNOWN_EVENT_IDS);
  }

  props.setProperty('knownEventIds', knownIds.join(','));
  return true;
}

// --- メインの実行関数 ---
function main(): void {
  try {
    console.log('Connpassイベント検索を開始します...');

    // スプレッドシートから設定を取得
    const config = getConfigFromSheet();
    console.log(`取得したキーワード数: ${config.keywords.length}`);

    let totalNewEvents = 0;

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

        // 新規イベントのみフィルタリング（既に過去1時間以内に更新されたイベントから）
        const newEvents = recentlyUpdatedEvents.filter(event =>
          isNewEvent(event.event_id)
        );

        if (newEvents.length > 0) {
          console.log(
            `"${keyword}" で ${newEvents.length}件の新規イベントが見つかりました`
          );

          // LINEでメッセージ送信
          const message = formatMultipleEvents(newEvents, keyword);
          sendLineMessage(message, config.lineChannelAccessToken);

          totalNewEvents += newEvents.length;
        } else {
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

    if (totalNewEvents === 0) {
      // 新規イベントがない場合はメッセージを送信せず、ログのみ出力
      const message =
        '過去1時間以内に更新された新しいイベントは見つかりませんでした。LINEメッセージは送信しません。';
      console.log(message);

      // スプレッドシートのUIにもメッセージを表示
      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('検索結果', message, ui.ButtonSet.OK);
      } catch (uiError) {
        console.log(
          'UI表示エラー（スクリプトエディタから実行された可能性があります）:',
          uiError
        );
      }
    } else {
      // 新規イベントがあった場合もUIに結果を表示
      const message = `処理完了: 合計 ${totalNewEvents}件の新規イベントを通知しました`;
      console.log(message);

      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('検索結果', message, ui.ButtonSet.OK);
      } catch (uiError) {
        console.log(
          'UI表示エラー（スクリプトエディタから実行された可能性があります）:',
          uiError
        );
      }
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

// --- テスト用関数 ---
function testConnection(): void {
  try {
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
    console.log(
      `LINE Channel Access Token value check:`,
      config.lineChannelAccessToken ? 'CUSTOM VALUE' : 'EMPTY'
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

    // テストメッセージを送信
    const testMessage =
      '🧪 GAS Connpassイベント検索アプリのテストメッセージです。';
    sendLineMessage(testMessage, config.lineChannelAccessToken);
    console.log('テストメッセージの送信に成功しました');

    // 実際の検索結果も送信（時間フィルタリングなし）
    const keyword = config.keywords[0];
    console.log(`テスト検索キーワード: "${keyword}"`);

    const events = searchConnpassEvents(keyword, config.connpassApiKey);
    console.log(`"${keyword}" で ${events.length}件のイベントが見つかりました`);

    if (events.length > 0) {
      // 時間フィルタリングなしで結果を送信
      let searchResultMessage = `🔍 「${keyword}」の検索結果 (全件: ${events.length}件)\n\n`;

      events.slice(0, MAX_EVENTS_PER_MESSAGE).forEach((event, index) => {
        searchResultMessage += `${index + 1}. ${formatEventMessage(event)}`;
        if (index < Math.min(events.length - 1, MAX_EVENTS_PER_MESSAGE - 1)) {
          searchResultMessage += '\n\n';
        }
      });

      if (events.length > MAX_EVENTS_PER_MESSAGE) {
        searchResultMessage += `\n\n他 ${events.length - MAX_EVENTS_PER_MESSAGE}件のイベントがあります。`;
      }

      sendLineMessage(searchResultMessage, config.lineChannelAccessToken);
      console.log('検索結果の送信に成功しました');
    } else {
      const noResultMessage = `🔍 「${keyword}」の検索結果: イベントが見つかりませんでした。`;
      sendLineMessage(noResultMessage, config.lineChannelAccessToken);
      console.log('検索結果なしメッセージの送信に成功しました');
    }

    // スプレッドシートのUIにも結果を表示
    try {
      const ui = SpreadsheetApp.getUi();
      const resultMessage = `接続テスト完了\nテストメッセージと検索結果（${events.length}件）を送信しました`;
      ui.alert('接続テスト結果', resultMessage, ui.ButtonSet.OK);
    } catch (uiError) {
      console.log(
        'UI表示エラー（スクリプトエディタから実行された可能性があります）:',
        uiError
      );
    }
  } catch (error) {
    console.error('接続テストでエラーが発生しました:', error);
    throw error;
  }
}

// --- スプレッドシート初期化関数 ---
function initializeSpreadsheet(): void {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // シートが空の場合のみ初期化
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === '') {
      initializeSpreadsheet();
      console.log('スプレッドシートの初期化が完了しました');
    }

    // カスタムメニューの追加
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Connpassイベント検索')
      .addItem('設定を初期化', 'initializeSpreadsheet')
      .addItem('検索結果送付テスト', 'testConnection')
      .addItem('検索結果フィルタリングテスト', 'main')
      .addToUi();
  } catch (error) {
    console.error('onOpen実行エラー:', error);
  }
}
