/**
 * ■ Gmail から Google カレンダーイベントを作成
 * ※このメソッドを時間主導型のトリガーに設定する
 * 
 * Gmail のラベルを活用して Times Car の予約関連メールを取得し、
 * その情報からカレンダーにイベントを新規登録・更新・削除できる。
 * ラベル名等の設定は Config.gs で管理している。
 */
function createEventFromGmail() {
  const conf = config(); // 定数の取得
  const timeZone = Session.getScriptTimeZone(); // スクリプトのタイムゾーン
  const mdl = new Model(); // Database（スプレッドシート）操作クラス
  const now = new Date();
  const nowStr = Utilities.formatDate(now, timeZone, 'yyyy/MM/dd HH:mm:ss');
  const calendar = !conf.calendarId ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(conf.calendarId); // カレンダーの取得

  // スプレッドシートから更新日時を取得
  const syncData = mdl.getData(conf.sheetNameSyncDateTime);
  const syncDateTime = syncData[0][conf.headerSyncDateTime];
  const syncDateStr = Utilities.formatDate(syncDateTime, timeZone, 'yyyy-MM-dd');
  const syncDateTimeStr = Utilities.formatDate(syncDateTime, timeZone, 'yyyy-MM-dd HH:mm:ss');

  // メールの取得（最終同期日以降、かつ TimesCar ラベルが付与されているもの）
  const queryStr = 'after:' + syncDateStr + ' label:' + conf.labelNameTimesCar;
  const threads = GmailApp.search(queryStr);
  const messages = threads.flatMap(thread => thread.getMessages());

  // 最終同期日時でメールをさらに絞り込み（GmailApp では時刻で絞り込めないため）
  const filteredMssgs = messages.filter(message => {
    const date = message.getDate();
    const target = new Date(syncDateTimeStr);
    return date >= target;
  });

  // 最新のメールが存在しなければ処理終了
  if (filteredMssgs.length === 0) return;

  // 各メッセージを処理
  let keyValuePair = {};
  let conditions = [];
  filteredMssgs.forEach(eachMssg => {
    // メールの内容から各項目の値を取得
    // 項目の値を取得
    const emailSubject = eachMssg.getSubject(); // メールの件名を取得
    const emailBody = eachMssg.getBody(); // メール本文のテキストを取得
    const reservationNumber = emailBody.match(conf.regExpRsvNumber) ? emailBody.match(conf.regExpRsvNumber)[1].trim() : null; // 予約番号
    const rsvStartDateTime = emailBody.match(conf.regExpRsvStartDateTime) ? new Date(emailBody.match(conf.regExpRsvStartDateTime)[1]) : null; // 予約:開始時刻の取得
    const rsvEndDatetime = emailBody.match(conf.regExpRsvEndDateTime) ? new Date(emailBody.match(conf.regExpRsvEndDateTime)[1]) : null; // 予約:終了時刻の取得
    const rtnStartDateTime = emailBody.match(conf.regExpRtnStartDateTime) ? new Date(emailBody.match(conf.regExpRtnStartDateTime)[1]) : null; // 返却:開始時刻の取得
    const rtnEndDatetime = emailBody.match(conf.regExpRtnEndDateTime) ? new Date(emailBody.match(conf.regExpRtnEndDateTime)[1]) : null; // 返却:終了時刻の取得

    // 同期履歴を参照
    conditions = [
      { key: conf.headerReservationNumber, value: Number(reservationNumber) }
    ];
    const syncHisData = mdl.getData(conf.sheetNameSyncHistory, conditions);

    // 履歴の有無と予約の状況で処理を分岐
    if (syncHisData.length === 0 && emailSubject.includes(conf.emailSubFactorRsvDone)) { // 履歴がなく、新規予約登録の場合
      // カレンダー登録
      const eventTitle = conf.subjectCalendarEvent + ' :' + reservationNumber;
      const event = calendar.createEvent(
        eventTitle,
        rsvStartDateTime,
        rsvEndDatetime,
        {
          description: emailBody
        }
      );
      event.setColor(CalendarApp.EventColor.YELLOW); // イベントの色を設定

      // Database データの格納
      const keyValuePairs = [{
        [conf.headerReservationNumber]: reservationNumber,
        [conf.headerEmailSubject]: emailSubject,
        [conf.headerLastStatus]: conf.emailSubFactorRsvDone,
        [conf.headerCalendarEventId]: event.getId(),
        [conf.headerCreatedDateTime]: nowStr,
        [conf.headerUpdatedDateTime]: nowStr
      }];
      Logger.log(mdl.insertData(conf.sheetNameSyncHistory, keyValuePairs));
    
    } else if (syncHisData.length > 0 && emailSubject.includes(conf.emailSubFactorRsvChanged)) { // 履歴があり、かつ予約情報の変更の場合
      // カレンダーの更新
      // イベントの取得
      const eventId = syncHisData[0][conf.headerCalendarEventId];
      const event = calendar.getEventById(eventId);
      // 情報の更新
      event.setTime(rsvStartDateTime, rsvEndDatetime); // 予約時刻
      event.setDescription(emailBody); // 説明欄

      // Database データの更新
      keyValuePair = {
        [conf.headerLastStatus]: conf.emailSubFactorRsvChanged,
        [conf.headerUpdatedDateTime]: nowStr
      };
      Logger.log(mdl.updateData(conf.sheetNameSyncHistory, keyValuePair, conditions));

    } else if (syncHisData.length > 0 && emailSubject.includes(conf.emailSubFactorRsvCanceled)) { // 履歴があり、かつ予約取り消しの場合
      // カレンダーの削除
      // イベントの取得
      const eventId = syncHisData[0][conf.headerCalendarEventId];
      const event = calendar.getEventById(eventId);
      // イベントの削除
      event.deleteEvent();

      // Database データの更新
      keyValuePair = {
        [conf.headerLastStatus]: conf.emailSubFactorRsvCanceled,
        [conf.headerUpdatedDateTime]: nowStr
      };
      Logger.log(mdl.updateData(conf.sheetNameSyncHistory, keyValuePair, conditions));

    } else if (syncHisData.length > 0 && emailSubject.includes(conf.emailSubFactorRsvAccomplished)) { // 履歴があり、かつ予約満了（返却）の場合
      // カレンダーの更新
      // イベントの取得
      const eventId = syncHisData[0][conf.headerCalendarEventId];
      const event = calendar.getEventById(eventId);
      // 情報の更新
      event.setTime(rtnStartDateTime, rtnEndDatetime); // 返却時刻
      event.setDescription(emailBody); // 説明欄

      // Database データの更新
      keyValuePair = {
        [conf.headerLastStatus]: conf.emailSubFactorRsvAccomplished,
        [conf.headerUpdatedDateTime]: nowStr
      };
      Logger.log(mdl.updateData(conf.sheetNameSyncHistory, keyValuePair, conditions));
    }
  });

  // Database の最終同期日時を更新
  keyValuePair = {
    [conf.headerSyncDateTime]: nowStr
  };
  conditions = [
    { key: '#', value: 1 }
  ];
  Logger.log(mdl.updateData(conf.sheetNameSyncDateTime, keyValuePair, conditions));
}
