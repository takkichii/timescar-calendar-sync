function config() {
  // スプレッドシート ID を指定しない（空欄）場合は getActiveSpreadsheet が実行される
  this.spreadsheetId = ""; // Spreadsheet ID: Default Database

  // ===========================
  // Setting as you like
  // ===========================
  this.settingAsYouLike = "Sample";

  this.labelNameTimesCar = 'TimesCar'; // Gmail の任意のラベル名
  this.subjectCalendarEvent = 'Times Car 予約'; // カレンダーのタイトル

  this.sheetNameSyncDateTime = '同期時刻管理';
  this.sheetNameSyncHistory = '同期履歴';

  this.headerSyncDateTime = '最終同期日時';
  this.headerReservationNumber = '予約番号';
  this.headerEmailSubject = 'メール件名';
  this.headerLastStatus = '最終ステータス';
  this.headerCalendarEventId = 'カレンダーイベントID';
  this.headerCreatedDateTime = '作成日';
  this.headerUpdatedDateTime = '更新日';

  this.titleReservationNumber = '■予約番号';
  this.titleRsvStartDateTime = '■利用開始日時';
  this.titleRsvEndDateTime = '■返却予定日時';

  this.emailSubFactorRsvDone = '予約登録完了';
  this.emailSubFactorRsvChanged = '予約変更完了';
  this.emailSubFactorRsvCanceled = '予約取消完了';
  this.emailSubFactorRsvAccomplished = '返却証';

  // 文字列取得のための補正値
  this.correctionStartDateTimeNum = 9;
  this.correctionEndDateTimeNum = 27;
  this.correctionStartRsvNum = 7;
  this.correctionEndRsvNum = 18;
  
  return this;
}