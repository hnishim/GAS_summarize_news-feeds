/**
 * メイン処理: スプレッドシートのRSSフィードを取得し、Geminiで要約してSlackに投稿
 */
function summarizeSpreadsheetWithGemini() {
  // --- 設定 ---
  const API_KEY_PROPERTY_NAME = 'GEMINI_API_KEY'; // APIキーを保存するプロパティ名
  const SLACK_WEBHOOK_URL_PROPERTY_NAME = 'SLACK_WEBHOOK_URL'; // Slack Webhook URL
  const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent'; // Gemini APIのエンドポイント
  const PROMPT_INSTRUCTION = "# 依頼内容\n以下は製薬会社等の記事URLです。URL先の内容を確認し、日本のAI創薬スタートアップを想定ユーザーとして、有用な情報提供をしてください。\n\n## 手順\n1. 以下の対象に合致するかを判断してください\n2. 条件に合致した場合：日本語で200字以内で、余計な枕詞は追加せず、記事の内容だけから要約を生成し、要約した文章のみを出力してください。\n3. 条件に合致しなかった場合：「Z」とだけ出力してください\n4. 対象の記事がないなど、上記の処理ができない場合：「Z」とだけ出力してください\n\n## 対象記事（例）\n- 創薬研究開発に関わる内容\n- 創薬研究開発におけるパートナーシップ、共同研究契約、産学連携\n\n## 対象外記事（例）\n以下のような内容は対象外とし、出力に含めないでください。\n- 決算発表\n- 創薬の内容を含まない中期経営計画や会社経営方針\n- 株式分割、配当金などのIR情報\n- Corporate Social Responsibility 活動\n- 創薬に関係のない記事（一般薬、ジェネリック医薬品、医療機器食品 等）\n\n# URL\n\n"; // 要約のための指示文
  const PROMPT_INSTRUCTION_EMAIL = "# 依頼内容\n以下は製薬会社等の記事です。内容を確認し、日本のAI創薬スタートアップを想定ユーザーとして、有用な情報提供をしてください。\n\n## 手順\n1. 以下の対象に合致するかを判断してください\n2. 条件に合致した場合：日本語で200字以内で、余計な枕詞は追加せず、記事の内容だけから要約を生成し、要約した文章のみを出力してください。\n3. 条件に合致しなかった場合：「Z」とだけ出力してください\n4. 対象の記事がないなど、上記の処理ができない場合：「Z」とだけ出力してください\n\n## 対象記事（例）\n- 創薬研究開発に関わる内容\n- 創薬研究開発におけるパートナーシップ、共同研究契約、産学連携\n\n## 対象外記事（例）\n以下のような内容は対象外とし、出力に含めないでください。\n- 決算発表\n- 創薬の内容を含まない中期経営計画や会社経営方針\n- 株式分割、配当金などのIR情報\n- Corporate Social Responsibility 活動\n- 創薬に関係のない記事（一般薬、ジェネリック医薬品、医療機器食品 等）\n\n# 記事\n\n"; // 要約のための指示文

  // --- rss_list_sheet シートの内容取得 ---
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rss_list_sheet = spreadsheet.getSheetByName("rss_list")
  let prompt_instruction = PROMPT_INSTRUCTION;

  Logger.log('スプレッドシートの内容を取得中...');

  var lastRow = rss_list_sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    const sheetName = rss_list_sheet.getRange(i, 2).getValue();
    const items_title = rss_list_sheet.getRange(i, 3);
    const items_summary = rss_list_sheet.getRange(i, 4);
    const items_url = rss_list_sheet.getRange(i, 5);
    const items_created = rss_list_sheet.getRange(i, 6);

    // すでに処理済みか確認
    const targetSheet = spreadsheet.getSheetByName(sheetName);
    Logger.log(sheetName);
    const lastRowInTargetSheet = targetSheet.getLastRow();
    const lastURLInTargetSheet = targetSheet.getRange(lastRowInTargetSheet, 3);

    if (items_url.getValue() == '#N/A' ) {
      Logger.log('フェードが取得されていないのでスキップします。');
      continue;
    }
    
    if ( items_url.getValue() == lastURLInTargetSheet.getValue() ) {
      Logger.log('処理済みのフィードのためスキップします。' + items_url.getValue());
      continue;
    };

    // 処理済みを記録するため、最終行にコピー
    const sheetId = spreadsheet.getSheetByName(sheetName).getSheetId();
    items_title.copyValuesToRange(sheetId, 1, 1, lastRowInTargetSheet + 1, lastRowInTargetSheet + 1);
    items_summary.copyValuesToRange(sheetId, 2, 2, lastRowInTargetSheet + 1, lastRowInTargetSheet + 1);
    items_url.copyValuesToRange(sheetId, 3, 3, lastRowInTargetSheet + 1, lastRowInTargetSheet + 1);
    items_created.copyValuesToRange(sheetId, 4, 4, lastRowInTargetSheet + 1, lastRowInTargetSheet + 1);

    // フィード要約
    prompt = prompt_instruction + items_url.getValue();

    Logger.log('スプレッドシートの内容取得完了。Geminiに送信します。');
    // Logger.log('送信内容のプレビュー（長すぎる場合は省略）: ' + allSheetsContent.substring(0, 1000) + '...'); // デバッグ用

    const pre_fix = "*" + sheetName + ",* " + items_created.getValue();
    const post_fix = items_url.getValue();
    summarizeWithGemini(prompt, pre_fix, post_fix);
  };

  summarizeEmailsWithGemini();

  // メールによるフィードの処理と削除
  function summarizeEmailsWithGemini() {
    const EMAIL_SHEET_TO_CLEAR = 'Email'; // 削除したいシート名
    const emailSheet = spreadsheet.getSheetByName(EMAIL_SHEET_TO_CLEAR);

    if (!emailSheet) {
        Logger.log(`注意: シート "${EMAIL_SHEET_TO_CLEAR}" が見つかりませんでした。`);
        return null;
    };

    Logger.log(`シート "${EMAIL_SHEET_TO_CLEAR}" の内容を処理します...`);
    var lastRow = emailSheet.getLastRow();
    for (var i = 1; i <= lastRow; i++) {
      const items_title = emailSheet.getRange(i, 1);
      const items_summary = emailSheet.getRange(i, 2);
      const items_from = emailSheet.getRange(i, 3);
      const items_created = emailSheet.getRange(i, 4);

      // フィード要約
      prompt = PROMPT_INSTRUCTION_EMAIL + items_title.getValue() + items_summary.getValue();

      Logger.log('スプレッドシートの内容取得完了。Geminiに送信します。');
      // Logger.log('送信内容のプレビュー（長すぎる場合は省略）: ' + allSheetsContent.substring(0, 1000) + '...'); // デバッグ用

      const pre_fix = '*' + items_from.getValue() + ',* ' + items_created.getValue();
      const post_fix = '';
      summarizeWithGemini(prompt, pre_fix, post_fix);
    };

    Logger.log(`シート "${EMAIL_SHEET_TO_CLEAR}" の内容を全削除します...`);
    emailSheet.clearContents();
    Logger.log(`シート "${EMAIL_SHEET_TO_CLEAR}" の内容を削除しました。`);
  };

  function summarizeWithGemini(prompt, pre_fix, post_fix) {
    // --- Gemini APIキーの取得 ---
    const apiKey = PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY_NAME);
    if (!apiKey) {
        Logger.log('エラー: APIキーが設定されていません。スクリプトプロパティに ' + API_KEY_PROPERTY_NAME + ' を設定してください。');
        Browser.msgBox('エラー: APIキーが設定されていません。スクリプトプロパティに ' + API_KEY_PROPERTY_NAME + ' を設定してください。');
        return;
    }

    // --- Gemini API リクエストの作成 ---
    const requestBody = {
        contents: [
          {
            parts: [
              {
                text: prompt
              }
            ]
          }
        ],
        // オプション（必要に応じて追加）：temperature, topK, topPなど
        generationConfig: {
         temperature: 0.5
        }
      };
  
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true // エラー時も例外を投げずに応答を取得
      };
  
      const apiUrlWithKey = `${GEMINI_API_URL}?key=${apiKey}`;
  
      // --- Gemini API の呼び出し ---
      try {
        Logger.log('Gemini APIを呼び出し中...');
        const response = UrlFetchApp.fetch(apiUrlWithKey, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
  
        Logger.log(`Gemini API レスポンスコード: ${responseCode}`);
        // Logger.log(`Gemini API レスポンスボディ: ${responseBody}`); // デバッグ時以外はコメントアウト推奨
  
        if (responseCode === 200) {
          const jsonResponse = JSON.parse(responseBody);
  
          // レスポンスから要約テキストを抽出
          // レスポンス構造はAPIのバージョンや応答内容によって変わる可能性あり
          if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
              jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
              jsonResponse.candidates[0].content.parts.length > 0 && jsonResponse.candidates[0].content.parts[0].text) {
  
            const summary = jsonResponse.candidates[0].content.parts[0].text; // 要約結果を取得
            Logger.log('要約が成功しました:\n' + summary);
  
            if (summary[0] == 'Z') {
              Logger.log('対象記事ではないため、ポストをスキップ');
              return null;
            };
  
            const postMessage = pre_fix + '\n' + summary + '\n' + post_fix;
  
            // Slack投稿処理
            sendSlackMessage(postMessage);
          } else {
            Logger.log('エラー: Geminiからの応答に要約テキストが見つかりませんでした。');
            Logger.log('応答ボディ: ' + responseBody);
            // エラーをSlackに通知することも可能
            sendSlackMessage('エラー: スプレッドシート要約中にGeminiからの応答に要約テキストが見つかりませんでした。ログを確認してください。');
          }
        } else {
          // APIエラーハンドリング
          Logger.log('エラー: Gemini APIの呼び出しに失敗しました。');
          Logger.log('応答コード: ' + responseCode);
          Logger.log('応答ボディ: ' + responseBody);
          // エラーをSlackに通知
          sendSlackMessage(`エラー: スプレッドシート要約中にGemini APIの呼び出しに失敗しました。ステータスコード: ${responseCode}。ログを確認してください。`);
        }
      } catch (e) {
        Logger.log('API呼び出し中に例外が発生しました: ' + e.toString());
        // 例外発生をSlackに通知
        sendSlackMessage(`エラー: スプレッドシート要約中に例外が発生しました。${e.toString()}。ログを確認してください。`);
      }  
  };
  
  /**
   * Slack Incoming Webhook を使ってメッセージを送信するヘルパー関数
   * @param {string} message - 送信するテキストメッセージ
   */
  function sendSlackMessage(message) {
    // --- Slack Webhook URLの取得 ---
    const slackWebhookUrl = PropertiesService.getScriptProperties().getProperty(SLACK_WEBHOOK_URL_PROPERTY_NAME);
    if (!slackWebhookUrl) {
        Logger.log('エラー: Slack Webhook URLが設定されていません。スクリプトプロパティに ' + SLACK_WEBHOOK_URL_PROPERTY_NAME + ' を設定してください。');
        Browser.msgBox('エラー: Slack Webhook URLが設定されていません...');
        return;
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        text: message // Slackに送るメッセージ本文
        // 他にも username, icon_emoji, attachments, blocks など設定可能
        // 例: username: "スプレッドシート要約Bot"
      }),
      muteHttpExceptions: true // エラー時も例外を投げない
    };

    try {
      Logger.log('Slackにメッセージを投稿中...');
      const response = UrlFetchApp.fetch(slackWebhookUrl, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        Logger.log('Slackへの投稿に成功しました。');
      } else {
        Logger.log('エラー: Slackへの投稿に失敗しました。');
        Logger.log('応答コード: ' + responseCode);
        Logger.log('応答ボディ: ' + response.getContentText());
      }
    } catch (e) {
      Logger.log('エラー: Slackへの投稿中に例外が発生しました: ' + e.toString());
    }
  };
};

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