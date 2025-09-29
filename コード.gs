// ==========================
// 定数・設定
// ==========================


// スプレッドシートのIDを定数として宣言
const SPREADSHEET_ID = "1Vphl8QIsrgKVr0_zribYDiU51jgfnkUZ_eMScem1E2M";

// 各シートのヘッダー名と関連情報を一元的に管理（JavaScriptのオブジェクトリテラル形式）
//経費区分申請の集計シート
const SHEET_CONFIGS = {
  DATA: {
    name: "form1",
    headers: {
      TIMESTAMP: "タイムスタンプ",
      CATEGORY: "経費区分",
      REMARKS: "備考",
      EMAIL: "メールアドレス",
    },
    format: {
      TIMESTAMP: "yyyy/MM/dd HH:mm:ss"
    }
  },
  //メンバーの名前とメールアドレス一覧のシート
  ROSTER: {
    name: "name",
    headers: {
      NAME: "名前",
      EMAIL: "アドレス",
      IS_AGGREGATION_TARGET: "集計対象"
    }
  },
  //イベントカレンダーに登録したデータ一覧シート
  EVENT: {
    name: "event",
    headers: {
      DATE: "日付",
      TITLE: "タイトル",
      MEMO: "メモ",
    },
    format: {
      DATE: "yyyy-MM-dd"
    }
  }
};

// ==========================
// ヘルパー関数
// ==========================

/**
 * ヘッダー名から列のインデックスを取得するヘルパー関数
 * @param {Array<string>} headers - ヘッダー名の配列
 * @param {string} targetHeader - 検索するヘッダー名
 * @returns {number} - ヘッダーのインデックス (見つからない場合は-1)
 */
function getColumnIndex(headers, targetHeader) {
  return headers.indexOf(targetHeader);
}

/**
 * ヘッダー名に対応する行の値を取得するヘルパー関数
 * @param {Array<any>} row - データ行の配列
 * @param {Array<string>} headers - ヘッダー名の配列
 * @param {string} targetHeader - 取得する値のヘッダー名
 * @returns {any} - 取得した値、または見つからない場合はnull
 */
function getValueByHeader(row, headers, targetHeader) {
  const colIndex = getColumnIndex(headers, targetHeader);
  if (colIndex === -1 || colIndex >= row.length) {
    return null; // ヘッダーが見つからない、またはインデックスが範囲外
  }
  const value = row[colIndex];
  return (value !== null && value !== undefined) ? value : ''; // null/undefinedを空文字列に変換
}

/**
 * スプレッドシートをIDで開き、見つからない場合はエラーを投げるヘルパー関数
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} - スプレッドシートオブジェクト
 */
function openSpreadsheetById() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) {
    throw new Error("指定されたスプレッドシートが見つからないか、アクセスできません。IDを確認してください。");
  }
  return ss;
}

/**
 * 現在のユーザーのメールアドレスを取得する
 * @returns {string} - ユーザーのメールアドレス
 */
function getUserEmail() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    // 開発環境などでメールアドレスが取得できない場合を考慮
    // 本番環境ではGoogleアカウントにログインしているユーザーのメールアドレスが取得されます
    return "anonymous@example.com";
  }
  return email;
}

/**
 * 名簿シートからメールアドレスと名前のマップを作成し、返す。
 * @returns {Object} メールアドレスをキー、名前を値とするオブジェクト
 */
function getEmailToRosterMap() {
  const ss = openSpreadsheetById();
  const rosterSheet = ss.getSheetByName(SHEET_CONFIGS.ROSTER.name); // 名簿シートを取得

  if (!rosterSheet) {
    throw new Error(`名簿シート '${SHEET_CONFIGS.ROSTER.name}' が見つかりません。`);
  }

  // 名簿シートのヘッダーを取得
  const rosterHeaders = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  // 名簿シートの値を2次元配列で取得
  const rosterData = rosterSheet.getDataRange().getValues();
  
  // CONFIG内の文言を変数に格納
  const rosterEmailHeader = SHEET_CONFIGS.ROSTER.headers.EMAIL;
  const rosterNameHeader = SHEET_CONFIGS.ROSTER.headers.NAME;

  // 名簿のシートのヘッダーで、メールアドレスと名前の列番号を変数に格納（見つからない場合は-1を返す）
  const emailColIndex = getColumnIndex(rosterHeaders, rosterEmailHeader);
  const nameColIndex = getColumnIndex(rosterHeaders, rosterNameHeader);

// 見つからない時（getColumnIndexが-1の時）の処理
  if (emailColIndex === -1 || nameColIndex === -1) {
    throw new Error(`名簿シートに'${rosterEmailHeader}' または '${rosterNameHeader}' ヘッダーが見つかりません。`);
  }

// 
  const emailToRosterMap = {};
  if (rosterData.length > 1) { // ヘッダー行を除外
    for (let i = 1; i < rosterData.length; i++) {
      // 2次元配列から1行ずつ取り出して処理
      const row = rosterData[i];
      // メールアドレスを小文字で取得
      const email = String(row[emailColIndex] || '').toLowerCase();
      const name = String(row[nameColIndex] || '');
      if (email) {
        emailToRosterMap[email] = name;
      }
    }
  }
  return emailToRosterMap;
}

// ==========================
// アクセス時のページ処理
// ==========================

/**
 * ウェブアプリとしてアクセスされたときに実行される関数。
 * アクセス制御を行い、許可されたユーザーには指定されたHTMLを表示します。
 * @param {GoogleAppsScript.Events.DoGet} e - GETリクエストのイベントオブジェクト
 * @returns {GoogleAppsScript.HTML.HtmlOutput} - 表示するHTMLコンテンツ
 */
function doGet(e) {
  // アクセスした人が、自分のログインしているGoogleアカウントの情報から、メールアドレスを小文字で取得し関数に格納
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  // getEmailToRosterMapというメールアドレスと名前を取得する関数を呼び、変数に格納
  const emailToRosterMap = getEmailToRosterMap();

// emailToRosterMap（名簿シートから取得した「メールアドレス→名前」のオブジェクト）に、
// 現在ログイン中のGoogleアカウント（Session.getActiveUser().getEmail()）が存在するかをチェックする。
  if (!emailToRosterMap.hasOwnProperty(userEmail)) {
    // アクセス権がない場合に、ブラウザに「アクセス拒否」ページを表示して処理を終了
    return HtmlService.createHtmlOutput(
      '<h1>アクセスが拒否されました。</h1>' +
      '<p>このページにアクセスする権限がありません。<br>' +
      '許可されたGoogleアカウントでログインしているかご確認ください。</p>'
    );
  }

  // URLパラメータ 'page' の値を取得
  const pageName = e.parameter.page;
  let template;

  // パラメータに応じて表示するテンプレートを選択
  // ここでどのページを返すかを明示的に制御する
  if (pageName === 'member') {
    template = HtmlService.createTemplateFromFile('member'); //メンバー管理用のページ「member.html」を指定
  }else if (pageName === 'knowledge') {
    template = HtmlService.createTemplateFromFile('knowledge');//ナレッジページの「Knowledge.html」を指定 
  } else {
    template = HtmlService.createTemplateFromFile('index'); // デフォルトは 「index.html」を指定
  }
  //「template」をevaluate()でHTML形式に変換する。
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');//レスポンシブルデザイン追加のためにmetaタグ追加
}

/**
 * Htmlファイルを読み込むためのヘルパー関数。
 * HTMLファイル内で `<?!= include('ファイル名'); ?>` を使用できるようにします。
 * @param {string} filename - 読み込むHTMLファイルのファイル名 (拡張子なし)
 * @returns {string} - ファイルの内容
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// ==========================
// 経費区分申請の処理
// ==========================

/**
 * HTMLから受け取った経費データをスプレッドシートに記録する。
 * @param {string} expenseType - 経費区分 ('通常立替', '客先立替', '立替なし')
 * @param {string} remarks - 備考
 * @returns {string} - 最新の集計データと経費一覧データ、または矛盾情報を含むJSON文字列
 */
function recordExpenseFromClick(expenseType, remarks) {
  const userEmail = getUserEmail();
  const today = new Date(); // 現在の日付と時刻を取得（Dateオブジェクトとして生成）
  const year = today.getFullYear(); // 現在の年（西暦）を取得する
  const month = today.getMonth() + 1; // 現在の月を取得する（0始まりのため +1 する）

  // 1. 今月は申請済みかチェックする
  const conflictingRecords = findConflictingRecords(userEmail, expenseType, year, month);

  // 2. 今月の申請が見つかった場合、上書き確認のメッセージを返す
  if (conflictingRecords.length > 0) { // 配列に要素があるかで判定
    // 申請済みの区分を取得
    const conflictingCategory = conflictingRecords[0].category; 
    //申請済みのタイムスタンプだけを抽出
    const conflictingTimestamps = conflictingRecords.map(r => r.timestamp); 
     // HTMLにJSON形式で返すメッセージを作成
    const message = `今月はすでに「${conflictingCategory}」で申請済みです。\n\n「${expenseType}」に上書きしますか？`;
    return JSON.stringify({
      status: 'conflict',
      message: message,
      timestampsToDelete: conflictingTimestamps // 削除対象のタイムスタンプの配列を返す
    });
  }

  // 3. 今月の申請がない場合、通常通り新しい申請を記録する
  const ss = openSpreadsheetById();
  const dataSheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);

  if (!dataSheet) {
    throw new Error(`シート '${SHEET_CONFIGS.DATA.name}' が見つかりません。`);
  }

  const headers = SHEET_CONFIGS.DATA.headers;
  const timestamp = new Date(); // 現在時刻をタイムスタンプとする

  // ヘッダーの順番に合わせてデータ配列を作成し、追加
  const headerOrder = [
    headers.TIMESTAMP,
    headers.CATEGORY,
    headers.REMARKS,
    headers.EMAIL
  ];

// ヘッダー順に並べた新しい行のデータ配列を作成
// switch文で各列に対応する値をセットし、想定外の列は空文字にする
  const newRowValues = []; //新しいレコードを格納する配列
  for (const headerKey of headerOrder) {
    switch (headerKey) {
      case headers.TIMESTAMP:
        newRowValues.push(timestamp);
        break;
      case headers.CATEGORY:
        newRowValues.push(expenseType);
        break;
      case headers.REMARKS:
        newRowValues.push(remarks);
        break;
      case headers.EMAIL:
        newRowValues.push(userEmail);
        break;
      default:
        newRowValues.push('');
        break;
    }
  }
 // 作成した新しい行のnewRowValuesをシートに追加する
  dataSheet.appendRow(newRowValues);

  // データ追加後、最新のデータと集計結果を返してHTML側を更新させる
  return getMonthlyDataAndSummary(timestamp.getFullYear(), timestamp.getMonth() + 1);
}

/**
 * 指定されたタイムスタンプの経費レコードを削除します。
 * @param {string} timestampToDeleteStr - 削除するタイムスタンプ文字列。
 * @return {object|null} 削除されたレコードのデータ、またはnull。
 */
function deleteExpenseRecord(timestampToDeleteStr) {
  Logger.log("--- deleteExpenseRecord デバッグログ開始 ---");
  Logger.log(`[JS->GAS] 渡されたタイムスタンプ文字列: '${timestampToDeleteStr}'`);

  if (!timestampToDeleteStr || timestampToDeleteStr === 'null') {
    Logger.log("GAS: `deleteExpenseRecord`関数に無効なタイムスタンプが渡されました。");
    Logger.log("--- deleteExpenseRecord デバッグログ終了 ---");
    return null;
  }

  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);
  
  if (!sheet || sheet.getLastRow() < 1) { // ★修正点：シートが存在しないか、データがない場合も考慮
    Logger.log(`シート '${SHEET_CONFIGS.DATA.name}' が見つからないか、データがありません。`);
    return null;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // ★修正点：headersが有効な配列であることを確認
  if (!Array.isArray(headers) || headers.length === 0) {
    Logger.log("ヘッダー行が取得できませんでした。");
    return null;
  }

  const timestampHeader = SHEET_CONFIGS.DATA.headers.TIMESTAMP;
  const timestampColIndex = getColumnIndex(headers, timestampHeader);
  if (timestampColIndex === -1) {
    throw new Error(`'${timestampHeader}' ヘッダーが見つかりません。`);
  }

  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const rowTimestampValue = row[timestampColIndex];

    if (rowTimestampValue instanceof Date) {
        const formattedRowTimestamp = Utilities.formatDate(
        rowTimestampValue,
        Session.getScriptTimeZone(),
        SHEET_CONFIGS.DATA.format.TIMESTAMP
        );

      if (formattedRowTimestamp === timestampToDeleteStr)  {
        sheet.deleteRow(i + 1);
        Logger.log(`レコードを削除しました。タイムスタンプ: ${timestampToDeleteStr}`);
        Logger.log("--- deleteExpenseRecord デバッグログ終了 ---");
        return {
          result: "success",
          deletedTimestamp: timestampToDeleteStr
        };
      }
    }
  }

  Logger.log(`指定されたタイムスタンプのレコードは見つかりませんでした: ${timestampToDeleteStr}`);
  Logger.log("--- deleteExpenseRecord デバッグログ終了 ---");
  return null;
}

/**
 * 指定されたタイムスタンプに対応する単一の経費レコードを取得します。
 * @param {string} timestampToFindStr - 検索するタイムスタンプ文字列。
 * @return {Object|null} 見つかったレコードをキーと値のペアのオブジェクトで返すか、見つからない場合はnullを返す。
 */
function getSingleExpenseRecord(timestampToFindStr) {
  Logger.log("--- getSingleExpenseRecord デバッグログ開始 ---");
  Logger.log(`[JS->GAS] 渡されたタイムスタンプ文字列: '${timestampToFindStr}'`);

  if (!timestampToFindStr || timestampToFindStr === 'null') { // 'null'文字列もチェック
    Logger.log("GAS: `getSingleExpenseRecord`関数に無効なタイムスタンプが渡されました。");
    Logger.log("--- getSingleExpenseRecord デバッグログ終了 ---");
    return null;
  }

  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);
  if (!sheet) {
    throw new Error(`シート '${SHEET_CONFIGS.DATA.name}' が見つかりません。`);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestampHeader = SHEET_CONFIGS.DATA.headers.TIMESTAMP;
  const timestampColIndex = getColumnIndex(headers, timestampHeader);
  if (timestampColIndex === -1) {
    throw new Error(`'${timestampHeader}' ヘッダーが見つかりません。`);
  }

  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const rowTimestampValue = row[timestampColIndex];

    if (rowTimestampValue instanceof Date) {
      // スプレッドシートのDateオブジェクトを、JavaScriptから渡された形式にフォーマット
      const formattedRowTimestamp = Utilities.formatDate(
        rowTimestampValue,
        Session.getScriptTimeZone(),
        SHEET_CONFIGS.DATA.format.TIMESTAMP
      );

      // フォーマット済み文字列同士で厳密に比較する
      if (formattedRowTimestamp === timestampToFindStr) {
        Logger.log(`一致しました！タイムスタンプ: ${timestampToFindStr}`);
        const record = {};
        headers.forEach((header, index) => {
          // `rowTimestampValue` ではなく `row` から値を取得
          let value = row[index];
          // タイムスタンプはフォーマット済み文字列に変換
          if (header === timestampHeader && value instanceof Date) {
            value = formattedRowTimestamp;
          }
          record[header] = value;
        });
        Logger.log("--- getSingleExpenseRecord デバッグログ終了 ---");
        return record;
      }
    }
  }

  Logger.log(`指定されたタイムスタンプのレコードは見つかりませんでした: ${timestampToFindStr}`);
  Logger.log("--- getSingleExpenseRecord デバッグログ終了 ---");
  return null;
}


/**
 * 経費レコードを更新する
 * @param {string} timestampToUpdate - 更新するレコードのタイムスタンプ
 * @param {Object} updatedData - 更新するデータを含むオブジェクト (例: {"経費区分": "新しい区分", "備考": "新しい備考"})
 */
function updateExpenseRecord(timestampToUpdate, updatedData) {
  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);
  if (!sheet) {
    throw new Error(`シート '${SHEET_CONFIGS.DATA.name}' が見つかりません。`);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestampColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.TIMESTAMP);

  if (timestampColIndex === -1) {
    throw new Error(`'${SHEET_CONFIGS.DATA.headers.TIMESTAMP}' ヘッダーが見つかりません。`);
  }

  const allData = sheet.getDataRange().getValues();
  let rowIndexToUpdate = -1;

  for (let i = 1; i < allData.length; i++) {
    const rowTimestampValue = allData[i][timestampColIndex];
    if (rowTimestampValue instanceof Date) {
      const formattedTimestamp = Utilities.formatDate(
        rowTimestampValue,
        Session.getScriptTimeZone(),
        SHEET_CONFIGS.DATA.format.TIMESTAMP
      );
      if (formattedTimestamp === timestampToUpdate) {
        rowIndexToUpdate = i;
        break;
      }
    }
  }

  if (rowIndexToUpdate === -1) {
    throw new Error(`タイムスタンプ '${timestampToUpdate}' のレコードが見つかりませんでした。`);
  }

  const rowData = allData[rowIndexToUpdate];
  for (const key in updatedData) {
    if (updatedData.hasOwnProperty(key)) {
      const colIndex = getColumnIndex(headers, key);
      if (colIndex !== -1) {
        rowData[colIndex] = updatedData[key];
      }
    }
  }
  // 更新された行をシートに書き戻す (GASの行番号は1ベースなので+1)
  sheet.getRange(rowIndexToUpdate + 1, 1, 1, rowData.length).setValues([rowData]);
  Logger.log(`Updated row: ${rowIndexToUpdate + 1} for timestamp: ${timestampToUpdate}`);
}

/**
 * 指定された年月の経費データと集計、未提出者リストを取得する
 * @param {number} year - 取得するデータの年
 * @param {number} month - 取得するデータの月 (1-12)
 * @returns {string} - 集計テキスト、経費一覧データ、未提出者リストを含むJSON文字列
 */
function getMonthlyDataAndSummary(year, month) {
  const ss = openSpreadsheetById();
  const dataSheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);
  const rosterSheet = ss.getSheetByName(SHEET_CONFIGS.ROSTER.name);

  // シートが見つからない場合のエラーチェックを追加
  if (!dataSheet || !rosterSheet) {
    throw new Error(`必要なシートが見つかりません。データシート名: ${SHEET_CONFIGS.DATA.name}, 名簿シート名: ${SHEET_CONFIGS.ROSTER.name}`);
  }

  // 経費データの取得と処理
  const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const rawData = dataSheet.getDataRange().getValues();

  // タイムスタンプ、経費区分、メールアドレスのヘッダー名を取得
  const timestampHeader = SHEET_CONFIGS.DATA.headers.TIMESTAMP;
  const categoryHeader = SHEET_CONFIGS.DATA.headers.CATEGORY;
  const emailHeader = SHEET_CONFIGS.DATA.headers.EMAIL;

  const timestampColIndex = getColumnIndex(dataHeaders, timestampHeader);
  const categoryColIndex = getColumnIndex(dataHeaders, categoryHeader);
  const emailColIndex = getColumnIndex(dataHeaders, emailHeader);

  // 経費データをフィルタリング
  const currentMonthData = []; // HTMLのテーブル表示用にすべての列を保持
  const submittedEmailsThisMonth = new Set(); // 今月提出した人のメールアドレス

  if (rawData.length > 1) { // ヘッダー行を除外
    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      const timestampValue = row[timestampColIndex];

      if (timestampValue instanceof Date) {
        const rowYear = timestampValue.getFullYear();
        const rowMonth = timestampValue.getMonth() + 1;

        if (rowYear === year && rowMonth === month) {
          // currentMonthData にはオブジェクト形式でデータを格納
          const rowObject = {};
          for (let j = 0; j < dataHeaders.length; j++) {
            let value = row[j];
            // タイムスタンプはフォーマットして格納
            if (dataHeaders[j] === timestampHeader && value instanceof Date) {
              value = Utilities.formatDate(value, Session.getScriptTimeZone(), SHEET_CONFIGS.DATA.format.TIMESTAMP);
            }
            rowObject[dataHeaders[j]] = value;
          }
          currentMonthData.push(rowObject);

          const email = row[emailColIndex];
          if (email) {
            submittedEmailsThisMonth.add(String(email).toLowerCase()); // メアドを小文字でセットに追加
          }
        }
      }
    }
  }

  // currentMonthData をタイムスタンプの新しい順にソートする
  currentMonthData.sort((a, b) => {
    // タイムスタンプはすでにフォーマット済み文字列 ("yyyy/M/d HH:mm:ss") になっているので、文字列比較で降順ソート
    const timestampA = a[timestampHeader];
    const timestampB = b[timestampHeader];
    return timestampB.localeCompare(timestampA);
  });

  // 名簿データの取得と処理
  const rosterHeaders = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  const rosterData = rosterSheet.getDataRange().getValues();

  const rosterNameHeader = SHEET_CONFIGS.ROSTER.headers.NAME;
  const rosterEmailHeader = SHEET_CONFIGS.ROSTER.headers.EMAIL;
  const rosterTargetHeader = SHEET_CONFIGS.ROSTER.headers.IS_AGGREGATION_TARGET;

  const rosterNameColIndex = getColumnIndex(rosterHeaders, rosterNameHeader);
  const rosterEmailColIndex = getColumnIndex(rosterHeaders, rosterEmailHeader);
  const rosterTargetColIndex = getColumnIndex(rosterHeaders, rosterTargetHeader);

  const allTargetMembers = []; // 集計対象の全メンバーのメールアドレス
  const allTargetMemberNames = new Map(); // メールアドレスと名前の紐付け

  if (rosterData.length > 1) { // ヘッダー行を除外
    for (let i = 1; i < rosterData.length; i++) {
      const row = rosterData[i];
      const isTarget = getValueByHeader(row, rosterHeaders, rosterTargetHeader);
      // '1' または 1 を集計対象とする
      if (String(isTarget).trim() === '1' || isTarget === 1) {
        const email = String(getValueByHeader(row, rosterHeaders, rosterEmailHeader)).toLowerCase();
        const name = String(getValueByHeader(row, rosterHeaders, rosterNameHeader));
        if (email) {
          allTargetMembers.push(email);
          allTargetMemberNames.set(email, name);
        }
      }
    }
  }

  // 未提出者リストの作成
  const unsubmittedList = [];
  for (const email of allTargetMembers) {
    if (!submittedEmailsThisMonth.has(email)) {
      const name = allTargetMemberNames.get(email);
      if (name) {
        unsubmittedList.push(name);
      }
    }
  }
  unsubmittedList.sort(); // 名前順にソート

  // 集計データの作成
  const categorizedNames = {};
  for (const rowObject of currentMonthData) {
    const categoriesStr = rowObject[categoryHeader] || '';
    const categories = categoriesStr.split(',').map(s => s.trim()).filter(s => s);
    const email = String(rowObject[emailHeader] || '').toLowerCase(); // メールアドレスを小文字で取得
    const name = allTargetMemberNames.get(email); // メールアドレスから名前を取得

    if (name) { // 名簿に存在する集計対象者のみ
      for (const cat of categories) {
        if (!categorizedNames[cat]) {
          categorizedNames[cat] = new Set();
        }
        categorizedNames[cat].add(name);
      }
    }
  }

  let summaryText = `今月の2solエンジニアメンバーの経費区分は以下になります。\n ※敬称略\n\n`;

  // カテゴリの固定順を定義
  const fixedCategoryOrder = ["通常立替", "客先立替", "立替なし"];

  // categorizedNames のキーを固定順でソートする
  const sortedCategories = fixedCategoryOrder.filter(category => categorizedNames.hasOwnProperty(category))
    .concat(
      Object.keys(categorizedNames)
      .filter(category => !fixedCategoryOrder.includes(category))
      .sort()
    );

  if (Object.keys(categorizedNames).length === 0) { // categorizedNames自体が空の場合
    summaryText += `${year}年${month}月にはデータがありません。`;
  } else {
    sortedCategories.forEach(category => {
      summaryText += `【${category}】\n`;
      const names = Array.from(categorizedNames[category]).sort();
      summaryText += (names.length === 0) ? `該当者なし\n` : names.join('\n') + '\n';
      summaryText += `\n`;
    });
  }

  // 結果をJSON形式で返す (このreturn文が関数の最後に来るように)
  return JSON.stringify({
    expenseData: {
      headers: dataHeaders, // テーブル表示用にヘッダーも返す
      data: currentMonthData // 今月の経費一覧データ
    },
    summaryText: summaryText,
    unsubmittedList: unsubmittedList
  });
}

/**
 * カレンダーイベントを読み込む
 */
function loadEvents() {
  const ss = openSpreadsheetById();
  const eventSheet = ss.getSheetByName(SHEET_CONFIGS.EVENT.name);
  if (!eventSheet) {
    throw new Error(`シート '${SHEET_CONFIGS.EVENT.name}' が見つかりません。`);
  }
  const eventHeaders = eventSheet.getRange(1, 1, 1, eventSheet.getLastColumn()).getValues()[0];
  const eventData = eventSheet.getDataRange().getValues();

  const events = [];
  if (eventData.length > 1) {
    const dateColIndex = getColumnIndex(eventHeaders, SHEET_CONFIGS.EVENT.headers.DATE);
    const titleColIndex = getColumnIndex(eventHeaders, SHEET_CONFIGS.EVENT.headers.TITLE);
    const memoColIndex = getColumnIndex(eventHeaders, SHEET_CONFIGS.EVENT.headers.MEMO);

    for (let i = 1; i < eventData.length; i++) {
      const row = eventData[i];
      const eventDate = row[dateColIndex];
      const eventTitle = row[titleColIndex];
      const eventMemo = row[memoColIndex];

      if (eventDate instanceof Date && eventTitle) {
        // フォーマット済みの日付文字列を返す
        const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), SHEET_CONFIGS.EVENT.format.DATE);
        events.push({
          "日付": formattedDate,
          "タイトル": eventTitle,
          "メモ": eventMemo || ""
        });
      }
    }
  }
  return events;
}

/**
 * 新しいカレンダーイベントを保存する
 * @param {string} date - イベント日付 (YYYY-MM-DD)
 * @param {string} title - イベントタイトル
 * @param {string} note - イベントメモ
 */
function saveEvent(date, title, note) {
  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.EVENT.name);
  if (!sheet) {
    throw new Error(`シート '${SHEET_CONFIGS.EVENT.name}' が見つかりません。`);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = getColumnIndex(headers, SHEET_CONFIGS.EVENT.headers.DATE);
  const titleColIndex = getColumnIndex(headers, SHEET_CONFIGS.EVENT.headers.TITLE);
  const memoColIndex = getColumnIndex(headers, SHEET_CONFIGS.EVENT.headers.MEMO);

  if (dateColIndex === -1 || titleColIndex === -1 || memoColIndex === -1) {
    throw new Error("必要なヘッダーが見つかりません。");
  }

  const newRow = new Array(headers.length).fill('');
  newRow[dateColIndex] = new Date(date);
  newRow[titleColIndex] = title;
  newRow[memoColIndex] = note;

  sheet.appendRow(newRow);

  try {
    const slackMessage = `新しいイベントが追加されました！\n\nタイトル: ${title}\n日付: ${date}\nメモ: ${note || "（なし）"}`;
    postToSlack(slackMessage, webhookUrl);
  } catch (e) {
    Logger.log("Slack通知の送信中にエラーが発生しました：" + e.message);
  }
}

/**
 * Slackにメッセージを投稿する
 * @param {string} message - 投稿するメッセージ
 * @param {string} webhookUrl - Slack Incoming Webhook URL
 */
function postToSlack(message, webhookUrl) {
  const payload = {
    text: message,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("Slackにメッセージを投稿しました。");
  } catch(e) {
    Logger.log(`Slackへの投稿に失敗しました: ${e.message}`);
  }
}

/**
 * カレンダーイベントを削除する
 * @param {string} dateToDelete - 削除するイベント日付 (YYYY-MM-DD)
 * @param {string} titleToDelete - 削除するイベントタイトル
 * @param {string} noteToDelete - 削除するイベントメモ (完全一致のため)
 */
function deleteEvent(dateToDelete, titleToDelete, noteToDelete) {
  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.EVENT.name);
  if (!sheet) {
    throw new Error(`シート '${SHEET_CONFIGS.EVENT.name}' が見つかりません。`);
  }
  
  // ★修正：スプレッドシートから直接ヘッダーを読み込む
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const dateColIndex = getColumnIndex(headers, SHEET_CONFIGS.EVENT.headers.DATE);
  const titleColIndex = getColumnIndex(headers, SHEET_CONFIGS.EVENT.headers.TITLE);
  const memoColIndex = getColumnIndex(headers, SHEET_CONFIGS.EVENT.headers.MEMO);

  // ヘッダーが正しく取得できたか確認
  if (dateColIndex === -1 || titleColIndex === -1 || memoColIndex === -1) {
    throw new Error("ヘッダーが見つからないか、不完全です。シートを確認してください。");
  }
  
  const data = sheet.getDataRange().getValues();
  // ヘッダー行を除外して検索、逆順に走査して削除時のインデックスずれを防ぐ
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const eventDate = row[dateColIndex];
    const eventTitle = row[titleColIndex];
    const eventMemo = row[memoColIndex];

    const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), SHEET_CONFIGS.EVENT.format.DATE);

    // 日付、タイトル、メモがすべて一致する行を削除
    if (formattedDate === dateToDelete && eventTitle === titleToDelete && (eventMemo || "") === (noteToDelete || "")) {
      sheet.deleteRow(i + 1); // スプレッドシートの行番号は1から始まる
      Logger.log(`Deleted event: Date=${dateToDelete}, Title=${titleToDelete}`);
      return;
    }
  }
  throw new Error(`指定されたイベントが見つかりませんでした。`);
}

/**
 * 既存のレコードを削除し、新しいレコードを記録する。
 * @param {string[]} timestampsToDelete 削除するレコードのタイムスタンプの配列
 * @param {string} expenseType 新しく記録する経費区分
 * @param {string} remarks 新しく記録する備考
 */
function deleteAndRecordExpense(timestampsToDelete, expenseType, remarks) {
  // まず、既存の矛盾する申請を削除
  if (timestampsToDelete && timestampsToDelete.length > 0) {
    for (const timestamp of timestampsToDelete) {
      deleteExpenseRecord(timestamp);
    }
  }

  // 次に、新しい申請を記録
  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);
  if (!sheet) {
    throw new Error(`シート '${SHEET_CONFIGS.DATA.name}' が見つかりません。`);
  }

  const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestampColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.TIMESTAMP);
  const categoryColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.CATEGORY);
  const remarksColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.REMARKS);
  const emailColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.EMAIL);

  const newRowData = new Array(headers.length).fill('');
  newRowData[timestampColIndex] = new Date();
  newRowData[categoryColIndex] = expenseType;
  newRowData[remarksColIndex] = remarks;
  newRowData[emailColIndex] = activeUserEmail;

  sheet.appendRow(newRowData);

  const today = new Date();
  return getMonthlyDataAndSummary(today.getFullYear(), today.getMonth() + 1);
}

/**
 * 矛盾する申請を探し、タイムスタンプとカテゴリを返すヘルパー関数。
 * @param {string} email 申請者のメールアドレス
 * @param {string} expenseType 申請しようとしている経費区分
 * @param {number} year 申請年
 * @param {number} month 申請月 (1-12)
 * @returns {Array<object>} 矛盾する申請のタイムスタンプとカテゴリの配列
 */
function findConflictingRecords(email, expenseType, year, month) {
  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.DATA.name);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestampColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.TIMESTAMP);
  const emailColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.EMAIL);
  const categoryColIndex = getColumnIndex(headers, SHEET_CONFIGS.DATA.headers.CATEGORY);

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const conflictingRecords = [];

  for (const row of values) {
    const recordTimestamp = row[timestampColIndex];
    const recordEmail = row[emailColIndex];
    const recordCategory = row[categoryColIndex];

    if (recordTimestamp instanceof Date) { // Dateオブジェクトであることを確認
      const recordDate = new Date(recordTimestamp);
      if (recordDate.getFullYear() === year && recordDate.getMonth() + 1 === month && recordEmail.toLowerCase() === email.toLowerCase()) {
        if ((recordCategory.includes('立替なし') && !expenseType.includes('立替なし')) ||
          (!recordCategory.includes('立替なし') && expenseType.includes('立替なし'))) {
          // タイムスタンプをフォーマットして返す
          const formattedTimestamp = Utilities.formatDate(
            recordTimestamp,
            Session.getScriptTimeZone(),
            SHEET_CONFIGS.DATA.format.TIMESTAMP
          );

          conflictingRecords.push({
            timestamp: formattedTimestamp,
            category: recordCategory
          });
        }
      }
    }
  }
  return conflictingRecords;
}

/**
 * 名簿メンバーリストを取得する
 */
function getMembers() {
  const ss = openSpreadsheetById();
  const sheet = ss.getSheetByName(SHEET_CONFIGS.ROSTER.name);
  if (!sheet) {
    throw new Error(`シート「${SHEET_CONFIGS.ROSTER.name}」が見つかりません。`);
  }
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];
  const nameCol = headers.indexOf(SHEET_CONFIGS.ROSTER.headers.NAME);
  const emailCol = headers.indexOf(SHEET_CONFIGS.ROSTER.headers.EMAIL);
  const isTargetCol = headers.indexOf(SHEET_CONFIGS.ROSTER.headers.IS_AGGREGATION_TARGET);

  if (nameCol === -1 || emailCol === -1 || isTargetCol === -1) {
    throw new Error('名簿シートのヘッダーが見つかりません。ヘッダーを確認してください。');
  }

  const members = values.slice(1).map(row => ({
    name: row[nameCol],
    email: row[emailCol],
    isTarget: (row[isTargetCol] === 1) }));

  return members;
}

/**
 * 名簿にメンバーを追加する
 */
function addMember(memberData) {
  const ss = openSpreadsheetById();
  const rosterSheet = ss.getSheetByName(SHEET_CONFIGS.ROSTER.name);
  if (!rosterSheet) {
    throw new Error('名簿シートが見つかりません。');
  }

  // ブール値を1または0に変換
  const isTargetValue = memberData.isTarget ? 1 : 0;

  // 書き込む行のデータを配列にまとめる
  const newRow = [
    memberData.name,
    memberData.email,
    isTargetValue // 変換した数値を入れる
  ];
  
  rosterSheet.appendRow(newRow);
  return 'メンバーが追加されました。';
}

/**
 * 名簿のメンバー情報を更新する
 */
function updateMember(rowIndex, memberData) {
  const ss = openSpreadsheetById();
  const rosterSheet = ss.getSheetByName(SHEET_CONFIGS.ROSTER.name);
  if (!rosterSheet) {
    throw new Error('名簿シートが見つかりません。');
  }
  
  // rowIndexは0始まりの配列のインデックスだが、スプレッドシートの行は1始まりなので+1する
  // ヘッダー行を考慮して、さらに+1する
  const sheetRowIndex = parseInt(rowIndex) + 2; 
  
  // ブール値を1または0に変換
  const isTargetValue = memberData.isTarget ? 1 : 0;

  // 書き込む範囲を指定
  const range = rosterSheet.getRange(sheetRowIndex, 1, 1, 3); // A列からC列の3列分
  
  // 書き込むデータを配列にまとめる
  const updatedRow = [
    memberData.name,
    memberData.email,
    isTargetValue // 変換した数値を入れる
  ];

  range.setValues([updatedRow]);
  return 'メンバー情報が更新されました。';
}

/**
 * 名簿のメンバーを削除する
 */
function deleteMember(rowIndex) {
  const ss = openSpreadsheetById(); 
  const sheet = ss.getSheetByName(SHEET_CONFIGS.ROSTER.name);
  if (!sheet) {
    throw new Error(`シート「${SHEET_CONFIGS.ROSTER.name}」が見つかりません。`);
  }
  const row = parseInt(rowIndex) + 2; // ヘッダー行と0-indexを考慮
  sheet.deleteRow(row);
}

const webhookUrl = "YOUR_SLACK_WEBHOOK_URL_HERE";

/**
 * Slackにメッセージを投稿する
 * @param {string} message - 投稿するメッセージ
 * @param {string} webhookUrl - Slack Incoming Webhook URL
 */
function postToSlack(message) {
  const payload = { text: message };
  const options = {
  method: "POST",
  contentType: "application/json",
  payload: JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("Slackにメッセージを投稿しました。");
  } catch(e) {
    Logger.log(`Slackへの投稿に失敗しました: ${e.message}`);
  }
}