// config.js
const CONFIG = {
  API: {
    GEMINI: {
      URL: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent',
      TEMPERATURE: 0.5
    }
  },
  PROPERTIES: {
    GEMINI_API_KEY: 'GEMINI_API_KEY',
    SLACK_WEBHOOK_URL: 'SLACK_WEBHOOK_URL'
  },
  SHEETS: {
    RSS_LIST: 'rss_list',
    EMAIL: 'Email'
  },
  PROMPTS: {
    RSS: `# 依頼内容
以下は製薬会社等の記事URLです。URL先の内容を確認し、日本のAI創薬スタートアップを想定ユーザーとして、有用な情報提供をしてください。

## 手順
1. 以下の対象に合致するかを判断してください
2. 条件に合致した場合：日本語で200字以内で、余計な枕詞は追加せず、記事の内容だけから要約を生成し、要約した文章のみを出力してください。
3. 条件に合致しなかった場合：「Z」とだけ出力してください
4. 対象の記事がないなど、上記の処理ができない場合：「Z」とだけ出力してください

## 対象記事（例）
- 創薬研究開発に関わる内容
- 創薬研究開発におけるパートナーシップ、共同研究契約、産学連携

## 対象外記事（例）
以下のような内容は対象外とし、出力に含めないでください。
- 決算発表
- 創薬の内容を含まない中期経営計画や会社経営方針
- 株式分割、配当金などのIR情報
- Corporate Social Responsibility 活動
- 創薬に関係のない記事（一般薬、ジェネリック医薬品、医療機器食品 等）

# URL\n\n`,
    EMAIL: `# 依頼内容
以下は製薬会社等の記事です。内容を確認し、日本のAI創薬スタートアップを想定ユーザーとして、有用な情報提供をしてください。

## 手順
1. 以下の対象に合致するかを判断してください
2. 条件に合致した場合：日本語で200字以内で、余計な枕詞は追加せず、記事の内容だけから要約を生成し、要約した文章のみを出力してください。
3. 条件に合致しなかった場合：「Z」とだけ出力してください
4. 対象の記事がないなど、上記の処理ができない場合：「Z」とだけ出力してください

## 対象記事（例）
- 創薬研究開発に関わる内容
- 創薬研究開発におけるパートナーシップ、共同研究契約、産学連携

## 対象外記事（例）
以下のような内容は対象外とし、出力に含めないでください。
- 決算発表
- 創薬の内容を含まない中期経営計画や会社経営方針
- 株式分割、配当金などのIR情報
- Corporate Social Responsibility 活動
- 創薬に関係のない記事（一般薬、ジェネリック医薬品、医療機器食品 等）

# 記事\n\n`
  }
};

// services/gemini.js
class GeminiService {
  constructor() {
    this.apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.GEMINI_API_KEY);
    if (!this.apiKey) {
      throw new Error(`APIキーが設定されていません。スクリプトプロパティに ${CONFIG.PROPERTIES.GEMINI_API_KEY} を設定してください。`);
    }
  }

  async summarize(prompt) {
    const requestBody = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: CONFIG.API.GEMINI.TEMPERATURE }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };

    const apiUrlWithKey = `${CONFIG.API.GEMINI.URL}?key=${this.apiKey}`;
    
    try {
      const response = UrlFetchApp.fetch(apiUrlWithKey, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        return this._extractSummary(jsonResponse);
      } else {
        throw new Error(`API呼び出しに失敗しました。ステータスコード: ${responseCode}`);
      }
    } catch (e) {
      throw new Error(`API呼び出し中に例外が発生しました: ${e.toString()}`);
    }
  }

  _extractSummary(jsonResponse) {
    if (jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text) {
      return jsonResponse.candidates[0].content.parts[0].text;
    }
    throw new Error('Geminiからの応答に要約テキストが見つかりませんでした。');
  }
}

// services/slack.js
class SlackService {
  constructor() {
    this.webhookUrl = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.SLACK_WEBHOOK_URL);
    if (!this.webhookUrl) {
      throw new Error(`Slack Webhook URLが設定されていません。スクリプトプロパティに ${CONFIG.PROPERTIES.SLACK_WEBHOOK_URL} を設定してください。`);
    }
  }

  async sendMessage(message) {
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(this.webhookUrl, options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error(`Slackへの投稿に失敗しました。ステータスコード: ${responseCode}`);
      }
    } catch (e) {
      throw new Error(`Slackへの投稿中に例外が発生しました: ${e.toString()}`);
    }
  }
}

// services/spreadsheet.js
class SpreadsheetService {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  getRssListSheet() {
    return this.spreadsheet.getSheetByName(CONFIG.SHEETS.RSS_LIST);
  }

  getEmailSheet() {
    return this.spreadsheet.getSheetByName(CONFIG.SHEETS.EMAIL);
  }

  copyRowToSheet(sourceRange, targetSheet, targetRow) {
    const sheetId = targetSheet.getSheetId();
    sourceRange.copyValuesToRange(sheetId, 1, 1, targetRow, targetRow);
  }
}

// main.js
function summarizeSpreadsheetWithGemini() {
  const geminiService = new GeminiService();
  const slackService = new SlackService();
  const spreadsheetService = new SpreadsheetService();

  try {
    // RSSフィードの処理
    processRssFeeds(geminiService, slackService, spreadsheetService);
    
    // メールの処理
    processEmails(geminiService, slackService, spreadsheetService);
  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.toString()}`);
  }

  function processRssFeeds(geminiService, slackService, spreadsheetService) {
    const rssListSheet = spreadsheetService.getRssListSheet();
    const lastRow = rssListSheet.getLastRow();
  
    for (let i = 2; i <= lastRow; i++) {
      const row = rssListSheet.getRange(i, 2, 1, 5);
      let [sheetName, title, summary, url, created] = row.getValues()[0];
  
      const rowToCopy = rssListSheet.getRange(i, 3, 1, 4);
  
      if (url === '#N/A') {
        Logger.log('フィードが取得されていないのでスキップします。');
        continue;
      }
  
      const targetSheet = spreadsheetService.spreadsheet.getSheetByName(sheetName);
      const lastUrlInTargetSheet = targetSheet.getRange(targetSheet.getLastRow(), 3).getValue();
  
      if (url === lastUrlInTargetSheet) {
        Logger.log('処理済みのフィードのためスキップします。' + url);
        continue;
      }
  
      if (sheetName == 'ミクスOnline') {
        url = 'https://www.mixonline.jp' + url;
      }
  
      // 新しい行を追加
      const newRow = targetSheet.getLastRow() + 1;
      spreadsheetService.copyRowToSheet(rowToCopy, targetSheet, newRow);
  
      // 要約処理
      const prompt = CONFIG.PROMPTS.RSS + url;
      const preFix = `*${sheetName},* ${created}`;
      const postFix = url;
  
      processSummary(geminiService, slackService, prompt, preFix, postFix);
    }
  }
  
  function processEmails(geminiService, slackService, spreadsheetService) {
    const emailSheet = spreadsheetService.getEmailSheet();
    if (!emailSheet) {
      Logger.log(`注意: シート "${CONFIG.SHEETS.EMAIL}" が見つかりませんでした。`);
      return;
    }
  
    const lastRow = emailSheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      const [title, summary, from, created] = emailSheet.getRange(i, 1, 1, 4).getValues()[0];
      
      const prompt = CONFIG.PROMPTS.EMAIL + title + summary;
      const preFix = `*${from},* ${created}`;
      const postFix = '';
  
      processSummary(geminiService, slackService, prompt, preFix, postFix);
    }
  
    // メールシートをクリア
    emailSheet.clearContents();
  }
  
  async function processSummary(geminiService, slackService, prompt, preFix, postFix) {
    try {
      const summary = await geminiService.summarize(prompt);
      
      if (summary[0] === 'Z') {
        Logger.log('対象記事ではないため、ポストをスキップ');
        return;
      }
  
      const message = `${preFix}\n${summary}\n${postFix}`;
      await slackService.sendMessage(message);
    } catch (e) {
      Logger.log(`要約処理中にエラーが発生しました: ${e.toString()}`);
      await slackService.sendMessage(`エラー: ${e.toString()}`);
    }
  }
}

/**
 * 特定の条件に一致するニュースレターメールをGmailから取得し、
 * スプレッドシートの指定シートに追記する。
 */
function processNewslettersFromEmails() {
  // --- 設定 ---
  const SHEET_NAME = 'Email'; // 追記したいシートの名前

  // Gmail検索クエリ: 取得したいメールを特定するクエリをここに記述 ★要変更★
  // 例: from:"newsletter@example.com" subject:"週刊 ニュース" is:unread
  // 例: label:"Newsletters/Important" is:unread
  const GMAIL_SEARCH_QUERY = 'is:unread label:"Newsfeed"'; // ★★★ここに適切な検索クエリを設定★★★
  const MAX_EMAILS_TO_PROCESS = 10; // 一度に処理するメールの最大数（多すぎると時間切れの可能性）

  Logger.log('ニュースレターの処理を開始します...');

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log(`エラー: シート "${SHEET_NAME}" が見つかりませんでした。シート名を確認してください。`);
      return;
    }

    // Gmailでメールを検索
    // is:unread を含めると未読メールのみが対象になります
    const threads = GmailApp.search(GMAIL_SEARCH_QUERY, 0, MAX_EMAILS_TO_PROCESS);

    if (threads.length === 0) {
      Logger.log('処理対象の未読メールは見つかりませんでした。');
      return;
    }

    Logger.log(`${threads.length} 件のスレッドが見つかりました。`);

    threads.forEach(thread => {
      // スレッド内の最新メッセージを取得（ニュースレターは通常これ）
      // または、thread.getMessages() ですべてのメッセージを取得
      const message = thread.getMessages().pop(); // 最新メッセージを取得

      // （もし検索クエリに is:unread を含めていない場合のみ）未読かどうかを確認
      // if (!message.isUnread()) {
      //   return; // 未読でない場合はスキップ
      // }

      try {
        const date = message.getDate();
        const subject = message.getSubject();
        const from = message.getFrom();
        const bodyHtml = message.getBody(); // HTML形式の本文
        const bodyPlain = message.getPlainBody(); // プレーンテキスト形式の本文

        Logger.log(`処理中: 件名 - "${subject}" (From: ${from}, Date: ${date})`);

        // ★★★ここからニュース本文の抽出ロジック★★★
        // 受け取るニュースレターの形式に合わせて、bodyHtml または bodyPlain から
        // 必要なニュース部分のテキストを抽出するコードを記述してください。
        // これは最もカスタマイズが必要な部分です。

        let extractedNewsContent = ""; // 抽出したニュース本文を格納する変数

        // 例1: プレーンテキスト本文をそのまま使う場合
        // extractedNewsContent = bodyPlain;

        // 例2: HTML本文から特定の要素（例: <div class="news-article">...</div>）の内容を抽出する場合
        // GASのXmlServiceを使ってHTMLをパースするのは少し複雑です。
        // 簡単な方法として、特定の開始マーカーと終了マーカー間のテキストを取得する方法があります（ただし正確性は低いです）。
        /*
        const startMarker = ""; // ニュース本文の開始を示すHTMLコメントなど
        const endMarker = "";   // ニュース本文の終了を示すHTMLコメントなど
        const startIndex = bodyHtml.indexOf(startMarker);
        const endIndex = bodyHtml.indexOf(endMarker);

        if (startIndex !== -1 && endIndex !== -1 && endIndex > startIndex) {
           extractedNewsContent = bodyHtml.substring(startIndex + startMarker.length, endIndex);
           // HTMLタグを取り除くなどの後処理が必要な場合が多い
           extractedNewsContent = extractedNewsContent.replace(/<[^>]*>/g, '').trim(); // 簡単なタグ除去
        } else {
           Logger.log('注意: 開始/終了マーカーが見つからなかったため、本文全体または一部を使用します。');
           extractedNewsContent = bodyPlain; // マーカーがない場合はプレーンテキスト全体を使うなど
        }
        */

        // 例3: プレーンテキスト本文から特定のキーワード間を抽出する場合
        /*
        const plainTextLines = bodyPlain.split('\n');
        let isExtracting = false;
        const extractedLines = [];
        for (const line of plainTextLines) {
            if (line.includes('--- ニュース本文開始 ---')) { // 開始キーワード
                isExtracting = true;
                continue; // 開始行自体は含めない場合
            }
            if (line.includes('--- ニュース本文終了 ---')) { // 終了キーワード
                isExtracting = false;
                break; // 終了行以降は処理しない
            }
            if (isExtracting) {
                extractedLines.push(line);
            }
        }
        extractedNewsContent = extractedLines.join('\n').trim();
        */

        // 上記の例を参考に、あなたのニュースレターの形式に合った抽出ロジックを記述してください。
        // ここでは例として、一旦プレーンテキスト本文をそのまま抽出結果とします。
        extractedNewsContent = bodyPlain; // ★★★ここをカスタマイズ★★★
        // ★★★抽出ロジックここまで★★★

        // スプレッドシートに追記するデータの配列
        const rowData = [
          subject,
          extractedNewsContent,
          from,
          date,
        ];

        // シートの最終行にデータを追記
        sheet.appendRow(rowData);
        Logger.log('スプレッドシートに行を追加しました。');

        // 処理済みとしてマーク
        message.markRead(); // メッセージを既読にする
      } catch (e) {
        Logger.log(`エラー: スレッド "${thread.getSubject()}" の処理中に例外が発生しました: ${e.toString()}`);
        // エラーが発生したメールを既読にせず残す、特定のエラーラベルを付けるなどの考慮も可能
      }
    });

     Logger.log('ニュースレターの処理が完了しました。');

  } catch (mainError) {
    Logger.log(`致命的なエラーが発生しました: ${mainError.toString()}`);
  }
}

/**
 * スプレッドシート上のすべてのシートに対し、A1セルにあると仮定した
 * IMPORTFEED関数を強制的に再計算させます。
 * 時間トリガーで使用することを想定しています。
 */
function refreshImportFeedFormulas() {
  // --- 設定 ---
  // このスクリプトが対象とするスプレッドシートのID。
  const SPREADSHEET_ID_FOR_REFRESH = SpreadsheetApp.getActiveSpreadsheet().getId();

  Logger.log(`全シートのA1セルのIMPORTFEED数式の更新を開始します (スプレッドシートID: ${SPREADSHEET_ID_FOR_REFRESH})...`);

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID_FOR_REFRESH);
    const sheets = spreadsheet.getSheets(); // スプレッドシート内のすべてのシートを取得

    let totalRefreshCount = 0; // 全シート合計の更新数

    if (sheets.length === 0) {
      Logger.log('スプレッドシートにシートが見つかりませんでした。');
      return;
    }

    Logger.log(`${sheets.length} 個のシートが見つかりました。各シートのA1セルを確認します。`);

    // 各シートをループ処理
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      Logger.log(`シート "${sheetName}" のA1セルを確認中...`);

      try {
        // A1セルを取得
        const cellA1 = sheet.getRange('A1');

        // A1セルの数式を取得
        const formulaA1 = cellA1.getFormula();

        // 数式があり、かつ IMPORTFEED 関数で始まるかチェック (大文字小文字を区別しない)
        if (formulaA1 && typeof formulaA1 === 'string' && formulaA1.toUpperCase().startsWith('=IMPORTFEED(')) {

          // IMPORTFEED関数がA1セルに見つかった場合
          Logger.log(`シート "${sheetName}" のA1セルにIMPORTFEED数式が見つかりました: ${formulaA1}`);

          // 同じ数式をA1セルに設定し直すことで、再計算を強制する
          cellA1.setFormula(formulaA1);
          totalRefreshCount++;
          Logger.log(`シート "${sheetName}" のA1セルの数式を再設定し、更新を試みました。`);

        } else {
          // A1セルに期待するIMPORTFEED関数がない場合
          Logger.log(`シート "${sheetName}" のA1セルはIMPORTFEED数式ではありません (${formulaA1 || '数式なし'})。このシートはスキップします。`);
        }
      } catch (sheetError) {
         // 特定のシートの処理中にエラーが発生した場合
         Logger.log(`エラー: シート "${sheetName}" の処理中に例外が発生しました: ${sheetError.toString()}`);
         // このシートはスキップし、次のシートの処理に進みます。
      }

      Logger.log(`シート "${sheetName}" の確認を完了しました。`);
    }); // sheets.forEach ループ終了

    Logger.log(`全シートのIMPORTFEED数式の更新処理が完了しました。合計 ${totalRefreshCount} 個の数式を再設定しました。`);

  } catch (mainError) {
    Logger.log(`IMPORTFEED数式更新中に致命的なエラーが発生しました: ${mainError.toString()}`);
  }
}