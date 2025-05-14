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
  EMAIL: {
    SEARCH_QUERY: 'is:unread label:"Newsfeed"',
    MAX_PROCESS_COUNT: 10
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

/**
 * スプレッドシートの内容を要約するクラス
 */
class SpreadsheetSummarizer {
  constructor() {
    this.geminiService = new GeminiService();
    this.slackService = new SlackService();
    this.spreadsheetService = new SpreadsheetService();
  }

  /**
   * スプレッドシートの内容を要約する
   */
  async summarize() {
    try {
      await this._processRssFeeds();
      await this._processEmails();
    } catch (error) {
      Logger.log(`エラーが発生しました: ${error.toString()}`);
      throw error;
    }
  }

  /**
   * RSSフィードを処理する
   */
  async _processRssFeeds() {
    const rssListSheet = this.spreadsheetService.getRssListSheet();
    const lastRow = rssListSheet.getLastRow();

    for (let i = 2; i <= lastRow; i++) {
      try {
        const row = rssListSheet.getRange(i, 2, 1, 5);
        const [sheetName, title, summary, url, created] = row.getValues()[0];

        if (url === '#N/A') {
          Logger.log('フィードが取得されていないのでスキップします。');
          continue;
        }

        const targetSheet = this.spreadsheetService.spreadsheet.getSheetByName(sheetName);
        const lastUrlInTargetSheet = targetSheet.getRange(targetSheet.getLastRow(), 3).getValue();

        if (url === lastUrlInTargetSheet) {
          Logger.log('処理済みのフィードのためスキップします。' + url);
          continue;
        }

        const processedUrl = this._processUrl(sheetName, url);
        await this._processFeed(sheetName, created, processedUrl, row);
      } catch (error) {
        Logger.log(`RSSフィードの処理中にエラー: ${error.toString()}`);
      }
    }
  }

  /**
   * URLを処理する
   * @param {string} sheetName シート名
   * @param {string} url URL
   * @returns {string} 処理後のURL
   */
  _processUrl(sheetName, url) {
    if (sheetName === 'ミクスOnline') {
      return 'https://www.mixonline.jp' + url;
    }
    return url;
  }

  /**
   * フィードを処理する
   * @param {string} sheetName シート名
   * @param {Date} created 作成日時
   * @param {string} url URL
   * @param {Range} row 行データ
   */
  async _processFeed(sheetName, created, url, row) {
    const targetSheet = this.spreadsheetService.spreadsheet.getSheetByName(sheetName);
    const newRow = targetSheet.getLastRow() + 1;
    const rowToCopy = row.getSheet().getRange(row.getRow(), 3, 1, 4);
    
    this.spreadsheetService.copyRowToSheet(rowToCopy, targetSheet, newRow);

    const prompt = CONFIG.PROMPTS.RSS + url;
    const preFix = `*${sheetName},* ${created}`;
    const postFix = url;

    await this._processSummary(prompt, preFix, postFix);
  }

  /**
   * メールを処理する
   */
  async _processEmails() {
    const emailSheet = this.spreadsheetService.getEmailSheet();
    if (!emailSheet) {
      Logger.log(`注意: シート "${CONFIG.SHEETS.EMAIL}" が見つかりませんでした。`);
      return;
    }

    const lastRow = emailSheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      try {
        const [title, summary, from, created] = emailSheet.getRange(i, 1, 1, 4).getValues()[0];
        
        const prompt = CONFIG.PROMPTS.EMAIL + title + summary;
        const preFix = `*${from},* ${created}`;
        const postFix = '';

        await this._processSummary(prompt, preFix, postFix);
      } catch (error) {
        Logger.log(`メールの処理中にエラー: ${error.toString()}`);
      }
    }

    emailSheet.clearContents();
  }

  /**
   * 要約を処理する
   * @param {string} prompt プロンプト
   * @param {string} preFix プレフィックス
   * @param {string} postFix ポストフィックス
   */
  async _processSummary(prompt, preFix, postFix) {
    try {
      const summary = await this.geminiService.summarize(prompt);
      
      if (summary[0] === 'Z') {
        Logger.log('対象記事ではないため、ポストをスキップ');
        return;
      }

      const message = `${preFix}\n${summary}\n${postFix}`;
      await this.slackService.sendMessage(message);
    } catch (error) {
      Logger.log(`要約処理中にエラー: ${error.toString()}`);
      await this.slackService.sendMessage(`エラー: ${error.toString()}`);
    }
  }
}

/**
 * ニュースレターメールを処理するクラス
 */
class EmailProcessor {
  constructor() {
    this.spreadsheetService = new SpreadsheetService();
  }

  /**
   * ニュースレターメールを処理する
   */
  async processNewsletters() {
    Logger.log('ニュースレターの処理を開始します...');

    try {
      const sheet = this.spreadsheetService.getEmailSheet();
      if (!sheet) {
        throw new Error(`シート "${CONFIG.SHEETS.EMAIL}" が見つかりませんでした。`);
      }

      const threads = this._searchEmailThreads();
      if (threads.length === 0) {
        Logger.log('処理対象の未読メールは見つかりませんでした。');
        return;
      }

      Logger.log(`${threads.length} 件のスレッドが見つかりました。`);
      await this._processThreads(threads, sheet);

      Logger.log('ニュースレターの処理が完了しました。');
    } catch (error) {
      Logger.log(`エラー: ${error.toString()}`);
      throw error;
    }
  }

  /**
   * メールスレッドを検索する
   * @returns {GmailThread[]} 検索結果のスレッド配列
   */
  _searchEmailThreads() {
    return GmailApp.search(
      CONFIG.EMAIL.SEARCH_QUERY,
      0,
      CONFIG.EMAIL.MAX_PROCESS_COUNT
    );
  }

  /**
   * スレッドを処理する
   * @param {GmailThread[]} threads 処理対象のスレッド配列
   * @param {Sheet} sheet 出力先シート
   */
  async _processThreads(threads, sheet) {
    for (const thread of threads) {
      try {
        const message = thread.getMessages().pop();
        const emailData = this._extractEmailData(message);
        
        if (emailData) {
          sheet.appendRow([
            emailData.subject,
            emailData.content,
            emailData.from,
            emailData.date
          ]);
          message.markRead();
          Logger.log(`メールを処理しました: ${emailData.subject}`);
        }
      } catch (error) {
        Logger.log(`スレッド "${thread.getSubject()}" の処理中にエラー: ${error.toString()}`);
      }
    }
  }

  /**
   * メールからデータを抽出する
   * @param {GmailMessage} message メールメッセージ
   * @returns {Object|null} 抽出したメールデータ
   */
  _extractEmailData(message) {
    try {
      return {
        date: message.getDate(),
        subject: message.getSubject(),
        from: message.getFrom(),
        content: message.getPlainBody()
      };
    } catch (error) {
      Logger.log(`メールデータの抽出中にエラー: ${error.toString()}`);
      return null;
    }
  }
}

/**
 * IMPORTFEED数式を更新するクラス
 */
class ImportFeedUpdater {
  constructor() {
    this.spreadsheetService = new SpreadsheetService();
  }

  /**
   * IMPORTFEED数式を更新する
   */
  async refreshAllSheets() {
    Logger.log('IMPORTFEED数式の更新を開始します...');

    try {
      const rssListSheet = this.spreadsheetService.getRssListSheet();
      if (!rssListSheet) {
        throw new Error(`シート "${CONFIG.SHEETS.RSS_LIST}" が見つかりませんでした。`);
      }

      const lastRow = rssListSheet.getLastRow();
      if (lastRow < 2) {
        Logger.log('更新対象の数式が見つかりませんでした。');
        return;
      }

      const targetRange = rssListSheet.getRange(2, 3, lastRow - 1, 4); // C2から最終行のF列まで
      const formulas = targetRange.getFormulas();
      let refreshCount = 0;

      for (let i = 0; i < formulas.length; i++) {
        for (let j = 0; j < formulas[i].length; j++) {
          const formula = formulas[i][j];
          if (this._isImportFeedFormula(formula)) {
            const cell = targetRange.getCell(i + 1, j + 1);
            cell.setFormula(formula);
            refreshCount++;
            Logger.log(`数式を更新しました: ${formula}`);
          }
        }
      }

      Logger.log(`更新完了: ${refreshCount} 個の数式を更新しました。`);
    } catch (error) {
      Logger.log(`エラー: ${error.toString()}`);
      throw error;
    }
  }

  /**
   * 数式がIMPORTFEED関数かどうかを判定する
   * @param {string} formula 数式
   * @returns {boolean} IMPORTFEED関数の場合はtrue
   */
  _isImportFeedFormula(formula) {
    return formula && 
           typeof formula === 'string' && 
           formula.toUpperCase().startsWith('=IMPORTFEED(');
  }
}

/**
 * スプレッドシートの内容を要約する
 */
function summarizeSpreadsheetWithGemini() {
  const summarizer = new SpreadsheetSummarizer();
  return summarizer.summarize();
}

/**
 * ニュースレターメールを処理する
 */
function processNewslettersFromEmails() {
  const processor = new EmailProcessor();
  return processor.processNewsletters();
}

/**
 * IMPORTFEED数式を更新する
 */
function refreshImportFeedFormulas() {
  const updater = new ImportFeedUpdater();
  return updater.refreshAllSheets();
}