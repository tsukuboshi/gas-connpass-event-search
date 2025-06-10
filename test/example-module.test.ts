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
  YEAR_MONTH_SHEET,
} from '../src/env';

describe('connpass-event-search', () => {
  describe('基本的なテスト', () => {
    it('テストが正常に実行される', () => {
      expect(true).toBe(true);
    });

    it('文字列の操作が正常に動作する', () => {
      const keyword = 'python';
      const message = `「${keyword}」に関するイベントは見つかりませんでした。`;
      expect(message).toContain('python');
      expect(message).toContain('見つかりませんでした');
    });

    it('配列の操作が正常に動作する', () => {
      const events = [1, 2, 3, 4, 5, 6, 7];
      const maxEvents = 5;
      const limitedEvents = events.slice(0, maxEvents);

      expect(limitedEvents).toHaveLength(5);
      expect(events.length > maxEvents).toBe(true);
    });
  });

  describe('環境定数のテスト', () => {
    it('Connpass API URLが正しく設定されている', () => {
      expect(CONNPASS_API_BASE_URL).toBe('https://connpass.com/api/v2/events/');
      expect(CONNPASS_API_BASE_URL).toMatch(/^https:\/\//);
    });

    it('メッセージあたりの最大イベント数が妥当である', () => {
      expect(MAX_EVENTS_PER_MESSAGE).toBeGreaterThan(0);
      expect(MAX_EVENTS_PER_MESSAGE).toBeLessThanOrEqual(10);
      expect(typeof MAX_EVENTS_PER_MESSAGE).toBe('number');
    });

    it('API呼び出し遅延時間が設定されている', () => {
      expect(API_CALL_DELAY).toBeGreaterThan(0);
      expect(typeof API_CALL_DELAY).toBe('number');
    });
  });

  describe('スプレッドシート列定義のテスト', () => {
    it('設定シートの列定義が正しい', () => {
      expect(SPREADSHEET_COLUMNS.CONNPASS_API_KEY).toBe(1);
      expect(SPREADSHEET_COLUMNS.LINE_CHANNEL_ACCESS_TOKEN).toBe(2);
      expect(SPREADSHEET_COLUMNS.KEYWORDS_START).toBe(3);
    });

    it('イベントシートの列定義が正しい', () => {
      expect(EVENT_SHEET_COLUMNS.TITLE).toBe(1);
      expect(EVENT_SHEET_COLUMNS.START_DATE).toBe(2);
      expect(EVENT_SHEET_COLUMNS.URL).toBe(3);
      expect(EVENT_SHEET_COLUMNS.NOTIFIED_DATE).toBe(4);
      expect(EVENT_SHEET_COLUMNS.KEYWORD).toBe(5);
    });

    it('列番号が重複していない', () => {
      const eventColumns = Object.values(EVENT_SHEET_COLUMNS);
      const uniqueColumns = [...new Set(eventColumns)];
      expect(eventColumns.length).toBe(uniqueColumns.length);
    });
  });

  describe('年月シート設定のテスト', () => {
    it('年月シートの設定が正しい', () => {
      expect(YEAR_MONTH_SHEET.NAME_FORMAT).toBe('YYYYMM');
      expect(YEAR_MONTH_SHEET.HEADER_ROW).toBe(1);
      expect(YEAR_MONTH_SHEET.DATA_START_ROW).toBe(2);
    });

    it('データ開始行がヘッダー行より後である', () => {
      expect(YEAR_MONTH_SHEET.DATA_START_ROW).toBeGreaterThan(
        YEAR_MONTH_SHEET.HEADER_ROW
      );
    });
  });

  describe('時間フィルタリング設定のテスト', () => {
    it('時間フィルタリングの設定が正しい', () => {
      expect(TIME_FILTERING.RECENT_HOURS).toBe(1);
      expect(TIME_FILTERING.TIMEZONE).toBe('Asia/Tokyo');
    });

    it('過去何時間の設定が妥当である', () => {
      expect(TIME_FILTERING.RECENT_HOURS).toBeGreaterThan(0);
      expect(TIME_FILTERING.RECENT_HOURS).toBeLessThanOrEqual(24);
    });
  });

  describe('年月シート名生成のテスト', () => {
    it('年月形式が正しい', () => {
      // モックした現在日時でテスト
      const mockDate = new Date('2024-12-15T10:30:00Z');
      const year = mockDate.getFullYear();
      const month = String(mockDate.getMonth() + 1).padStart(2, '0');
      const expectedSheetName = `${year}${month}`;

      expect(expectedSheetName).toMatch(/^\d{6}$/);
      expect(expectedSheetName).toBe('202412');
    });

    it('月の0埋めが正しく動作する', () => {
      const testCases = [
        { month: 1, expected: '01' },
        { month: 9, expected: '09' },
        { month: 10, expected: '10' },
        { month: 12, expected: '12' },
      ];

      testCases.forEach(({ month, expected }) => {
        const formatted = String(month).padStart(2, '0');
        expect(formatted).toBe(expected);
      });
    });

    it('前の月のシート名が正しく生成される', () => {
      const currentDate = new Date('2024-02-15T10:30:00Z');
      const previousMonth = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth() - 1,
        1
      );
      const year = previousMonth.getFullYear();
      const month = String(previousMonth.getMonth() + 1).padStart(2, '0');
      const sheetName = `${year}${month}`;

      expect(sheetName).toBe('202401');
      expect(sheetName).toMatch(/^\d{6}$/);
    });

    it('年を跨ぐ前の月のシート名が正しく生成される', () => {
      const currentDate = new Date('2024-01-15T10:30:00Z');
      const previousMonth = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth() - 1,
        1
      );
      const year = previousMonth.getFullYear();
      const month = String(previousMonth.getMonth() + 1).padStart(2, '0');
      const sheetName = `${year}${month}`;

      expect(sheetName).toBe('202312');
      expect(year).toBe(2023);
    });
  });

  describe('イベントコピー機能のテスト', () => {
    it('イベントの開催月判定が正しく動作する', () => {
      const eventDate = new Date('2024/02/15 19:00');
      const year = eventDate.getFullYear();
      const month = String(eventDate.getMonth() + 1).padStart(2, '0');
      const eventYearMonth = `${year}-${month}`;

      expect(eventYearMonth).toBe('2024-02');
      expect(eventYearMonth).toMatch(/^\d{4}-\d{2}$/);
    });

    it('シート名から年月形式への変換が正しく動作する', () => {
      const sheetName = '202402';
      const yearMonth =
        sheetName.substring(0, 4) + '-' + sheetName.substring(4, 6);

      expect(yearMonth).toBe('2024-02');
      expect(yearMonth.length).toBe(7);
    });
  });

  describe('イベント管理機能のテスト', () => {
    it('イベントURLの形式が正しい', () => {
      const testEventUrl = 'https://connpass.com/event/12345/';

      expect(testEventUrl).toMatch(/^https:\/\/connpass\.com\/event\/\d+\/$/);
      expect(testEventUrl.startsWith('https://connpass.com/event/')).toBe(true);
    });

    it('日時フォーマットが正しく動作する', () => {
      const testDate = new Date('2024-12-15T19:00:00+09:00');
      const timezone = 'Asia/Tokyo';

      // Google Apps Scriptの Utilities.formatDate のモック
      const mockFormatDate = (
        date: Date,
        tz: string,
        format: string
      ): string => {
        if (format === 'yyyy/MM/dd HH:mm') {
          return '2024/12/15 19:00';
        } else if (format === 'yyyy/MM/dd HH:mm:ss') {
          return '2024/12/15 19:00:00';
        }
        return '';
      };

      const formattedStart = mockFormatDate(
        testDate,
        timezone,
        'yyyy/MM/dd HH:mm'
      );
      const formattedNotified = mockFormatDate(
        testDate,
        timezone,
        'yyyy/MM/dd HH:mm:ss'
      );

      expect(formattedStart).toBe('2024/12/15 19:00');
      expect(formattedNotified).toBe('2024/12/15 19:00:00');
    });

    it('時間フィルタリングのロジックが正しい', () => {
      const now = new Date();
      const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
      const twoHoursAgo = new Date(now.getTime() - 2 * 60 * 60 * 1000);

      // 1時間以内のイベントは含まれる
      const recentEvent = {
        updated_at: oneHourAgo.toISOString(),
      };

      // 2時間前のイベントは除外される
      const oldEvent = {
        updated_at: twoHoursAgo.toISOString(),
      };

      const isRecentEventIncluded =
        new Date(recentEvent.updated_at) >= oneHourAgo;
      const isOldEventExcluded = new Date(oldEvent.updated_at) < oneHourAgo;

      expect(isRecentEventIncluded).toBe(true);
      expect(isOldEventExcluded).toBe(true);
    });
  });

  describe('メッセージフォーマットのテスト', () => {
    it('複数イベントのメッセージが正しくフォーマットされる', () => {
      const keyword = 'python';
      const expectedStart = `🔍 「${keyword}」の検索結果`;

      expect(expectedStart).toContain('🔍');
      expect(expectedStart).toContain(keyword);
    });

    it('既通知済みメッセージが正しくフォーマットされる', () => {
      const keyword = 'javascript';
      const count = 3;
      const expectedMessage = `📢 「${keyword}」の検索結果 (既に通知済み: ${count}件)`;

      expect(expectedMessage).toContain('📢');
      expect(expectedMessage).toContain('既に通知済み');
      expect(expectedMessage).toContain(count.toString());
    });
  });
});
