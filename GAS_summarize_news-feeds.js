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
    BASE_PROMPT: `# 依頼内容
以下は製薬会社等の記事です。内容を確認し、日本のAI創薬スタートアップを想定ユーザーとして、有用な情報提供をしてください。

## 手順
1. 以下の対象に合致するかを判断してください
2. 条件に合致した場合：日本語で200字以内で、余計な枕詞は追加せず、記事の内容だけから要約を生成し、要約した文章のみを出力してください。
3. 条件に合致しなかった場合：「Z」とだけ出力してください
4. 対象の記事がないなど、上記の処理ができない場合：「Z」とだけ出力してください

## 対象記事（例）
- 創薬研究開発に関わる内容
- 創薬研究開発におけるパートナーシップ、共同研究契約、産学連携
- 治験開始、医薬品の承認申請・承認、医薬品の適用拡大

## 対象外記事（例）
以下のような内容は対象外とし、出力に含めないでください。
- 決算発表
- 創薬の内容を含まない中期経営計画や会社経営方針
- 株式分割、配当金などのIR情報
- Corporate Social Responsibility 活動
- 創薬に関係のない記事（一般薬、ジェネリック医薬品、医療機器食品 等）`,

    RSS: (content, url) => `${CONFIG.PROMPTS.BASE_PROMPT}

# 記事
${content}

# URL
${url}`,

    EMAIL: (content) => `${CONFIG.PROMPTS.BASE_PROMPT}

# 記事
${content}`
  }
};

// services/gemini.js
class GeminiService {
  constructor() {
    this.apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.GEMINI_API_KEY);
    if (!this.apiKey) {
      throw new Error(`APIキーが設定されていません。スクリプトプロパティに ${CONFIG.PROPERTIES.GEMINI_API_KEY} を設定してください。`);
    }
    this.maxRetries = 3; // 最大リトライ回数
    this.retryDelay = 60000; // リトライ間隔（ミリ秒）
  }

  async summarize(prompt) {
    let lastError = null;
    
    for (let attempt = 1; attempt <= this.maxRetries; attempt++) {
      try {
        const requestBody = {
          contents: [{ parts: [{ text: prompt }] }],
          tools: [{googleSearch: {}}],
          generationConfig: { temperature: CONFIG.API.GEMINI.TEMPERATURE }
        };

        const options = {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(requestBody),
          muteHttpExceptions: true
        };

        const apiUrlWithKey = `${CONFIG.API.GEMINI.URL}?key=${this.apiKey}`;
        
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
        lastError = e;
        if (attempt < this.maxRetries) {
          Logger.log(`API呼び出し中にエラーが発生しました（試行 ${attempt}/${this.maxRetries}）: ${e.toString()}`);
          Logger.log(`${this.retryDelay/1000}秒後に再試行します...`);
          await Utilities.sleep(this.retryDelay);
        }
      }
    }

    // すべてのリトライが失敗した場合
    throw new Error(`API呼び出し中に例外が発生しました（${this.maxRetries}回試行）: ${lastError.toString()}`);
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
   * エラー値をnullに置き換える
   * @param {*} value チェックする値
   * @returns {*} エラー値の場合はnull、それ以外は元の値
   */
  _replaceErrorValue(value) {
    if (value === null || value === undefined) {
      return null;
    }

    // 文字列としてのエラー値のチェック
    if (typeof value === 'string') {
      const errorValues = ['#N/A', '#ERROR!', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'];
      if (errorValues.includes(value)) {
        return null;
      }
    }

    // スプレッドシートのエラーオブジェクトのチェック
    if (value instanceof Error) {
      return null;
    }

    return value;
  }

  /**
   * 日付をYYYY/M/d形式にパースする
   * @param {*} dateValue パースする日付値
   * @returns {string|null} パースされた日付文字列（YYYY/M/d）またはnull
   */
  _parseDate(dateValue) {
    // エラー値の場合はnullを返す
    if (this._replaceErrorValue(dateValue) === null) {
      return null;
    }

    try {
      let date;
      
      // すでにDate型の場合はそのまま使用
      if (dateValue instanceof Date) {
        date = dateValue;
      }
      // 文字列の場合はパースを試みる
      else if (typeof dateValue === 'string') {
        // スラッシュ区切りの日付（YYYY/MM/dd, MM/dd/YYYY, dd/MM/YYYY）
        if (dateValue.includes('/')) {
          const parts = dateValue.split('/');
          if (parts.length === 3) {
            // YYYY/MM/dd形式
            if (parts[0].length === 4) {
              date = new Date(parts[0], parts[1] - 1, parts[2]);
            }
            // MM/dd/YYYY形式またはdd/MM/YYYY形式
            else if (parts[2].length === 4) {
              const month = parseInt(parts[0], 10);
              const day = parseInt(parts[1], 10);
              
              // 月が1-12の範囲内かチェック
              if (month >= 1 && month <= 12) {
                // 日が1-31の範囲内かチェック
                if (day >= 1 && day <= 31) {
                  // MM/dd/YYYY形式として解釈
                  date = new Date(parts[2], month - 1, day);
                } else {
                  // 日が範囲外の場合はdd/MM/YYYY形式として解釈
                  date = new Date(parts[2], parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
                }
              } else {
                // 月が範囲外の場合はdd/MM/YYYY形式として解釈
                date = new Date(parts[2], parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
              }
            }
          }
        }
        // ハイフン区切りの日付（YYYY-MM-dd）
        else if (dateValue.includes('-')) {
          date = new Date(dateValue);
        }
        // その他の形式はDate.parse()で試行
        else {
          date = new Date(Date.parse(dateValue));
        }
      }
      // 数値の場合はタイムスタンプとして解釈
      else if (typeof dateValue === 'number') {
        date = new Date(dateValue);
      }
      // その他の型の場合はnullを返す
      else {
        return null;
      }

      // 日付が有効かチェック
      if (isNaN(date.getTime())) {
        return null;
      }

      // YYYY/M/d形式にフォーマット（先頭の0を省略）
      const year = date.getFullYear();
      const month = date.getMonth() + 1;  // 0を省略
      const day = date.getDate();         // 0を省略
      return `${year}/${month}/${day}`;
    } catch (error) {
      Logger.log(`日付のパース中にエラー: ${error.toString()}`);
      return null;
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

        // エラー値のチェックと置換
        const processedTitle = this._replaceErrorValue(title);
        const processedSummary = this._replaceErrorValue(summary);
        const processedUrl = this._processUrl(sheetName, this._replaceErrorValue(url));
        const processedDate = this._parseDate(created);

        // title, summary, url のすべてがnullの場合はスキップ
        if (processedTitle === null && processedSummary === null && processedUrl === null) {
          Logger.log('フィードの必須情報が取得されていないのでスキップします。');
          continue;
        }

        const targetSheet = this.spreadsheetService.spreadsheet.getSheetByName(sheetName);
        const lastUrlInTargetSheet = targetSheet.getRange(targetSheet.getLastRow(), 3).getValue();

        if (processedUrl === lastUrlInTargetSheet) {
          Logger.log('処理済みのフィードのためスキップします。' + processedUrl);
          continue;
        }

        await this._processFeed(sheetName, processedTitle, processedSummary, processedUrl, processedDate, row);
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
   * RSSフィードを処理する
   * @param {string} sheetName シート名
   * @param {string} title 記事タイトル
   * @param {string} summary 記事サマリー
   * @param {string} url 記事URL
   * @param {string} created 作成日時（YYYY/MM/dd形式）
   * @param {Range} row 行データ
   */
  async _processFeed(sheetName, title, summary, url, created, row) {
    const targetSheet = this.spreadsheetService.spreadsheet.getSheetByName(sheetName);
    const newRow = targetSheet.getLastRow() + 1;
    const rowToCopy = row.getSheet().getRange(row.getRow(), 3, 1, 4);
    
    this.spreadsheetService.copyRowToSheet(rowToCopy, targetSheet, newRow);

    // nullの場合は空文字列に置換
    const safeTitle = title || '';
    const safeSummary = summary || '';
    const safeDate = created || '';
    const prompt = CONFIG.PROMPTS.RSS(safeTitle + `\n` + safeSummary, url);
    const preFix = `*${sheetName},* ${safeDate}`;
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
        
        const prompt = CONFIG.PROMPTS.EMAIL(title + `\n` + summary);
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
          // メール処理後に少し待機
          await Utilities.sleep(100);
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
            // 数式の更新後に少し待機
            await Utilities.sleep(100);
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

async function summarizeSpreadsheetWithGemini() {
  const slackService = new SlackService(); // Slack通知用にインスタンスを作成

  // IMPORTFEED数式を更新
  try {
    const updater = new ImportFeedUpdater();
    await updater.refreshAllSheets();
  } catch (error) {
    const errorMessage = `IMPORTFEED数式の更新中にエラーが発生しました: ${error.toString()}`;
    Logger.log(errorMessage);
    await slackService.sendMessage(errorMessage); // Slackにエラー通知
  }
  
  // ニュースレターメールをシートに転記
  try {
    const processor = new EmailProcessor();
    await processor.processNewsletters();
  } catch (error) {
    const errorMessage = `ニュースレターメールの処理中にエラーが発生しました: ${error.toString()}`;
    Logger.log(errorMessage);
    await slackService.sendMessage(errorMessage); // Slackにエラー通知
  }

  // スプレッドシートの内容を要約
  try {
    const summarizer = new SpreadsheetSummarizer();
    await summarizer.summarize();
  } catch (error) {
    Logger.log(`スプレッドシートの要約処理中にエラーが発生しました: ${error.toString()}`);
    throw error; // エラーを再スローして、GASの実行ログに記録されるようにする
  }
}