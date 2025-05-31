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
});
