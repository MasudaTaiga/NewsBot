const APP_CONFIG = {
  SEARCH: {
    SPREADSHEET_ID: "",
    SHEET_NAME: "設定",
    KEYWORD_COLUMN: 1,
    PERIOD: "1d",
    MAX_RSS_ITEMS: 10,
    DISPLAY_LIMIT: 15
  },
  
  WEBHOOK: {
    URL: "",
    TITLE: "ITニュースクリップ"
  },

  UI: {
    TAG_COLOR: "#5f6368",
    DATE_COLOR: "#9aa0a6"
  }
};

function sendDailyItNewsToChat() {
  try {
    const allNewsMap = fetchAndScoreNews();
    const sortedNews = sortNews(allNewsMap);
    if (sortedNews.length === 0) {
      console.log("配信対象のニュースが見つかりませんでした。");
      return;
    }
    const payload = buildChatPayload(sortedNews);
    postToGoogleChat(payload);
  } catch (error) {
    console.error(`実行エラー: ${error.message}`);
  }
}

function fetchAndScoreNews() {
  const { SPREADSHEET_ID, SHEET_NAME, KEYWORD_COLUMN, PERIOD, MAX_RSS_ITEMS } = APP_CONFIG.SEARCH;
  const newsMap = new Map();
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return newsMap;

  const keywords = sheet.getRange(2, KEYWORD_COLUMN, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(k => k && k.toString().trim() !== "");

  keywords.forEach(keyword => {
    const url = `https://news.google.com/rss/search?q=${encodeURIComponent(keyword)}+when:${PERIOD}&hl=ja&gl=JP&ceid=JP:ja`;
    try {
      const response = UrlFetchApp.fetch(url);
      const xml = XmlService.parse(response.getContentText());
      const items = xml.getRootElement().getChild('channel').getChildren('item').slice(0, MAX_RSS_ITEMS);

      items.forEach(item => {
        const link = item.getChildText('link');
        const title = item.getChildText('title').replace(/["'<>]/g, "");
        const pubDate = item.getChildText('pubDate');
        const formattedDate = pubDate ? Utilities.formatDate(new Date(pubDate), "JST", "MM/dd HH:mm") : "";

        if (newsMap.has(link)) {
          const data = newsMap.get(link);
          if (!data.hitKeywords.includes(keyword)) {
            data.hitKeywords.push(keyword);
            data.score += 1;
          }
        } else {
          newsMap.set(link, {
            title, link, date: formattedDate, hitKeywords: [keyword], score: 1
          });
        }
      });
    } catch (e) {
      console.warn(`キーワード「${keyword}」の取得に失敗: ${e.message}`);
    }
  });
  return newsMap;
}

function sortNews(newsMap) {
  return Array.from(newsMap.values())
    .sort((a, b) => b.score - a.score)
    .slice(0, APP_CONFIG.SEARCH.DISPLAY_LIMIT);
}

function buildChatPayload(newsList) {
  const { SPREADSHEET_ID, PERIOD } = APP_CONFIG.SEARCH;
  const { TITLE } = APP_CONFIG.WEBHOOK;
  const { TAG_COLOR, DATE_COLOR } = APP_CONFIG.UI;

  const today = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  const dynamicSubtitle = `${today} 抽出 [範囲: ${PERIOD}]`;
  const ssUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
  
  const infoWidget = {
    textParagraph: {
      text: `⚙️ <a href="${ssUrl}">スプレッドシート</a>のキーワードに基づき抽出しています`
    }
  };

  const newsWidgets = newsList.map(news => {
    const tags = news.hitKeywords.map(k => `#${k}`).join(" ");
    return {
      textParagraph: {
        text: `<font color="${TAG_COLOR}"><small>${tags}</small></font><br>` + 
              `<a href="${news.link}"><b>${news.title}</b></a><br>` + 
              `<font color="${DATE_COLOR}"><small>${news.date} (Score: ${news.score})</small></font>`
      }
    };
  });

  return {
    cardsV2: [{
      cardId: "itNewsCard",
      card: {
        header: {
          title: TITLE,
          subtitle: dynamicSubtitle
        },
        sections: [{
          widgets: [infoWidget, ...newsWidgets]
        }]
      }
    }]
  };
}

function postToGoogleChat(payload) {
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(APP_CONFIG.WEBHOOK.URL, options);
  if (response.getResponseCode() !== 200) {
    console.error(`Webhook送信エラー: ${response.getContentText()}`);
  } else {
    console.log("Google Chatへの送信が完了しました。");
  }
}
