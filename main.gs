const Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("connpass");
const ClmIndex = {
  EventId: 1,
  Title: 2,
  StartedAt: 3,
  EventUrl: 4,
  CrawlDate: 5,
};

// スプレッドシートから取得
const LINE_ACCESS_TOKEN = Sheet.getRange("LINE_ACCESS_TOKEN").getValue();
const USER_ID = Sheet.getRange("LINE_USER_ID").getValue();
const CONNPASS_GROUP_ID = Sheet.getRange("CONNPASS_GROUP_ID").getValue();

const LINE_COLOR_CODE = "#06c755";
const LINE_PUSH_REQUEST = "https://api.line.me/v2/bot/message/push";
const CAROUSEL_TITLE_MAX_LENGTH = 40;
const CONNPASS_REQUEST = `https://connpass.com/api/v1/event/?series_id=${CONNPASS_GROUP_ID}&count=3&order2`;

// https://connpass.com/about/api/

/**
 * GASで設定するトリガー関数
 */
const getLineDCEvents = () => {
  const response = UrlFetchApp.fetch(CONNPASS_REQUEST);
  const json = JSON.parse(response.getContentText());

  // ISO8601 形式で現在時刻を取得
  const isoNow = getNowDateIso();

  // 終了イベントを削除
  deletePassedEvents(isoNow);

  // 明日開催イベントを通知
  notifyTomorrowEvents(isoNow);

  // 新規登録イベントを通知
  notifyNewEvents(json.events, isoNow);
};

const getNowDateIso = () => {
  const now = new Date();
  now.setHours(now.getHours() + 9);

  // "2022-11-18T19:00:00+09:00";//
  return new Date(now).toISOString().split(".")[0] + "+09:00";
};

/**
 * 過去のイベントをシートから削除
 */
const deletePassedEvents = (isoNow) => {
  const events = getEvents();

  const deleteTargetEventRowIds = [];
  const currentMsec = new Date(isoNow).getTime();

  events.forEach((e, index) => {
    const targetMsec = new Date(e[ClmIndex.StartedAt - 1]).getTime();
    if (targetMsec < currentMsec) {
      const currentRowId = index + 2;
      deleteTargetEventRowIds.push(currentRowId);
    }
  });

  if (deleteTargetEventRowIds.length === 0) {
    // 削除対象がなければ何もしない
    return;
  }

  deleteTargetEventRowIds.reverse().forEach((id) => {
    // 下の行から削除
    Sheet.deleteRow(id);
  });
};

/**
 * 翌日開催のイベントを通知
 */
const notifyTomorrowEvents = (isoNow) => {
  const events = getEvents();

  const today = new Date(isoNow);
  today.setDate(today.getDate() + 1);

  const tomorrow = Utilities.formatDate(new Date(today), "GMT+9", "yyyy/MM/dd");
  const carouselContents = [];

  events.forEach((e, index) => {
    const eventDate = Utilities.formatDate(
      new Date(e[ClmIndex.StartedAt - 1]),
      "GMT+9",
      "yyyy/MM/dd"
    );

    if (eventDate !== tomorrow) {
      return;
    }

    // イベントURLをもとにHTMLを解析しogImageを取得。カルーセルのサムネイルに使っている
    const ogImage = getOgImage(e[ClmIndex.EventUrl - 1]);
    const date = Utilities.formatDate(
      new Date(e[ClmIndex.StartedAt - 1]),
      "GMT+9",
      "yyyy/MM/dd hh:mm:ss"
    );

    const param = {
      thumbnailImageUrl: ogImage,
      imageBackgroundColor: LINE_COLOR_CODE,
      title: e[ClmIndex.Title - 1].substr(0, CAROUSEL_TITLE_MAX_LENGTH),
      text: `日時 ${date}`,
      actions: [
        {
          type: "uri",
          label: "イベントページへ",
          uri: e[ClmIndex.EventUrl - 1],
        },
      ],
    };

    carouselContents.push(param);
  });

  if (carouselContents.length === 0) {
    // 明日開催のイベントがない場合
    return;
  }

  // 明日のイベントをPush
  pushToLine("明日開催のイベントがあります", carouselContents);
};

/**
 * 新着イベントを通知
 */
const notifyNewEvents = (connpassEvents, isoNow) => {
  const events = getEvents();
  const eventList = events.map((e, index) => {
    return {
      eventId: e[ClmIndex.EventId - 1],
      rowId: index + 2,
    };
  });

  const filteredConnpassEvents = [];
  const currentMsec = new Date(isoNow).getTime();

  // 既にシートに登録されているイベント or 過去のイベントは除外
  connpassEvents.forEach((connpassEvent) => {
    const targetMsec = new Date(connpassEvent.started_at).getTime();
    if (
      eventList.map((e) => e.eventId).indexOf(connpassEvent.event_id) === -1 &&
      targetMsec > currentMsec
    ) {
      filteredConnpassEvents.push(connpassEvent);
    }
  });

  // 新着イベントなしの場合は通知しない
  if (filteredConnpassEvents.length === 0) {
    return;
  }

  let pushMessage = "新着イベントです\n\n";
  const carouselContents = [];

  // 未登録のイベントをシートに追加
  filteredConnpassEvents.reverse().forEach((e) => {
    Sheet.appendRow([e.event_id, e.title, e.started_at, e.event_url, isoNow]);
    const date = Utilities.formatDate(
      new Date(e.started_at),
      "GMT+9",
      "yyyy/MM/dd hh:mm"
    );
    pushMessage += `${e.title}\n日時：${date}\n${e.event_url}`;

    const ogImage = getOgImage(e.event_url);
    const param = {
      thumbnailImageUrl: ogImage,
      imageBackgroundColor: LINE_COLOR_CODE,
      title: e.title.substr(0, CAROUSEL_TITLE_MAX_LENGTH),
      text: `日時 ${date}`,
      actions: [
        {
          type: "uri",
          label: "イベントページへ",
          uri: e.event_url,
        },
      ],
    };

    carouselContents.push(param);
  });

  // 未登録イベントをPush
  pushToLine("Everything It's you!! 新規イベントのお知らせ", carouselContents);
};

/**
 * スプレッドシートからイベント情報部分をまるっと取得
 */
const getEvents = () => {
  const events = Sheet.getDataRange().getValues();

  // 先頭は項目名なので削除
  events.splice(0, 1);

  return events;
};

/**
 * LINEにプッシュメッセージ送信
 */
const pushToLine = (messageTitle, carouselContents) => {
  const pushMessage = {
    to: USER_ID,
    messages: [
      {
        type: "text",
        text: messageTitle,
      },
      {
        type: "template",
        altText: "connpass通知くん",
        template: {
          type: "carousel",
          columns: carouselContents,
        },
      },
    ],
  };

  const param = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${LINE_ACCESS_TOKEN}`,
    },
    payload: JSON.stringify(pushMessage),
  };

  try {
    UrlFetchApp.fetch(LINE_PUSH_REQUEST, param);
  } catch (e) {
    console.log(e);
  }
};
