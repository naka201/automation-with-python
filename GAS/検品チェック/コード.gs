// ====================================================================
// SECTION 0: 全システム共通の設定項目
// ====================================================================

// 1. 「検品管理データベース」スプレッドシートのID
// (Webアプリと日次レポートの両方で使用します)
//const DATABASE_SHEET_ID = '127UO42QG5yuYqg41IV5bjL-jmNA4Z9hW3Uiu47Vn5vI'; 
const DATABASE_SHEET_ID = '<database_id>'; 

// 2. 「入力済みマスター」スプレッドシートのID
// (日次レポート機能でのみ使用します)
//const INPUT_SHEET_ID = '1_RpyL2sYRR7WyuwTpXbjaw9phCOAYKROTCrxu2Td-xw';
const INPUT_SHEET_ID = '<input_id>';

// 3. 結果ファイルを出力するGoogleドライブのフォルダID  
// (日次レポート機能でのみ使用します)
const OUTPUT_FOLDER_ID = '1joEREmy4QHc_7tZ_acjUSM5tjCQUd--o';
const OUTPUT_FOLDER_ID = '<output_id>';


// --- グローバル変数 ---
const SPREADSHEET = SpreadsheetApp.openById(DATABASE_SHEET_ID);
const INSPECTION_SHEET = SPREADSHEET.getSheetByName('検品シート');
const ACCOUNT_SHEET = SPREADSHEET.getSheetByName('アカウント');
const BOX_NUMBER_SHEET = SPREADSHEET.getSheetByName('箱番シート');


// ====================================================================
// SECTION 1: Webアプリ関連の関数 (ログイン、検品処理など)
// ====================================================================

/**
 * ウェブアプリのUIを表示するためのメイン関数
 */
function doGet(e) {
  const userProperties = PropertiesService.getUserProperties();
  const isLoggedIn = userProperties.getProperty('loggedIn');

  if (isLoggedIn === 'true') {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('検品確認システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createTemplateFromFile('login.html').evaluate()
      .setTitle('ログイン')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * ログイン状態をチェックする関数
 */
function checkLogin(username, password) {
  const accounts = getAccountInfo();
  if (accounts[username] && accounts[username] === password) {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('loggedIn', 'true');
    userProperties.setProperty('username', username);
    return true;
  }
  return false;
}

/**
 * ログアウト処理
 */
function logout() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('loggedIn');
  userProperties.deleteProperty('username');
}

/**
 * アカウント情報をスプレッドシートから取得する関数
 */
function getAccountInfo() {
  const data = ACCOUNT_SHEET.getDataRange().getValues();
  const accounts = {};
  for (let i = 1; i < data.length; i++) {
    accounts[data[i][0]] = data[i][1]; // A列がユーザー名, B列がパスワード
  }
  return accounts;
}

/**
 * HTMLテンプレートに他のHTMLファイルをインクルードするための関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 箱番号をキーに検品シートから情報を検索する関数
 */
function getBoxInfo(boxNumber) {
  try {
    const data = INSPECTION_SHEET.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() == boxNumber) { // A列が箱番
        return {
          row: i + 1,
          vendorName: data[i][1],
          inspected: data[i][2],
          missing: data[i][3]
        };
      }
    }
    return null;
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * 欠番処理を行う関数
 */
function processMissing(boxNumber) {
  const boxInfo = getBoxInfo(boxNumber);
  if (!boxInfo || boxInfo.error) return 'エラー: 箱番号が見つかりません。';
  
  const now = new Date();
  INSPECTION_SHEET.getRange(boxInfo.row, 4).check(); // D列（欠番）
  INSPECTION_SHEET.getRange(boxInfo.row, 5).setValue(now); // E列にタイムスタンプ
  INSPECTION_SHEET.getRange(boxInfo.row, 6).clearContent(); // F列（点数）
  
  return `箱番 ${boxNumber} を欠番処理しました。`;
}

/**
 * 検品完了処理を行う関数
 */
function processInspectionComplete(data) {
  const { boxNumber, score, fileData, fileName, mimeType } = data;
  if (parseInt(score, 10) > 10) return 'エラー: 点数は10以下の数値を入力してください。';
  
  const boxInfo = getBoxInfo(boxNumber);
  if (!boxInfo || boxInfo.error) return 'エラー: 箱番号が見つかりません。';
  
  try {
    const decodedData = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);
    
    const now = new Date();
    const folderPath = `/検品チェック/画像/${Utilities.formatDate(now, "Asia/Tokyo", "yyyy_MM")}`;
    const folder = getOrCreateFolder(folderPath);

    const fileUrl = folder.createFile(blob).getUrl();
    const username = PropertiesService.getUserProperties().getProperty('username');
    
    INSPECTION_SHEET.getRange(boxInfo.row, 3).check(); // C列（検品済み）
    INSPECTION_SHEET.getRange(boxInfo.row, 5).setValue(now); // E列にタイムスタンプ
    INSPECTION_SHEET.getRange(boxInfo.row, 6).setValue(score); // F列に点数
    INSPECTION_SHEET.getRange(boxInfo.row, 7).setValue(fileUrl); // G列に画像リンク
    INSPECTION_SHEET.getRange(boxInfo.row, 8).setValue(username); // H列にアカウント名
    
    return `箱番 ${boxNumber} の検品を完了しました。`;
  } catch (e) {
    return `エラーが発生しました: ${e.message}`;
  }
}

/**
 * 指定されたパスのフォルダを取得または作成するヘルパー関数
 */
function getOrCreateFolder(path) {
  let parentFolder = DriveApp.getRootFolder();
  path.split('/').filter(Boolean).forEach(folderName => {
    const folders = parentFolder.getFoldersByName(folderName);
    parentFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  });
  return parentFolder;
}

/**
 * 月初のトリガーで実行し、その月の画像を保存するフォルダを作成する関数
 */
function createMonthlyFolder() {
  const now = new Date();
  const folderPath = `/検品チェック/画像/${Utilities.formatDate(now, "Asia/Tokyo", "yyyy_MM")}`;
  getOrCreateFolder(folderPath);
  Logger.log(`Successfully created or found folder: ${folderPath}`);
}


// ====================================================================
// SECTION 2: 日次進捗レポート機能 (トリガーで実行)
// ====================================================================

/**
 * メイン関数：この関数を日次トリガーで実行する
 */
function runDailyProgressCheck() {
  try {
    const today = new Date();
    
    const enteredDataSummary = aggregateEnteredData();
    const inspectedDataSummary = fetchAndFilterInspectedData(today);

    // アカウントごとに箱数と合計点数を集計
    const accountProgress = new Map();
    let totalScoreToday = 0;
    inspectedDataSummary.todaysInspections.forEach(item => {
      const progress = accountProgress.get(item.account) || { boxCount: 0, totalScore: 0 };
      progress.boxCount++;
      progress.totalScore += item.score;
      accountProgress.set(item.account, progress);
      totalScoreToday += item.score; // 本日の総点数も集計
    });

    const todaysUnenteredBoxes = inspectedDataSummary.todaysInspectedBoxSet.filter(box => 
      !enteredDataSummary.allEnteredBoxSet.has(box)
    );
    const recoveredBoxes = checkRecoveredBoxes(enteredDataSummary.allEnteredBoxSet);
    const vendorProgress = calculateVendorProgress(enteredDataSummary, inspectedDataSummary.allInspectedData);

    generateReport(today, inspectedDataSummary, enteredDataSummary, todaysUnenteredBoxes, recoveredBoxes, vendorProgress, accountProgress, totalScoreToday);

    console.log('本日の進捗チェック処理が正常に完了しました。');
  } catch(e) {
    console.error(`エラーが発生しました: ${e.stack}`);
  }
}

/**
 * 1. 「入力済み」シートを解析し、業者ごとの入力状況を集計する
 */
function aggregateEnteredData() {
  const sheet = SpreadsheetApp.openById(INPUT_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  const vendorData = new Map();
  const allEnteredBoxSet = new Set();

  for (let i = 1; i < data.length; i++) { // 2行目から
    const row = data[i];
    const vendorCode = row[0];
    if (!vendorCode) continue;

    const totalBoxes = parseInt(row[3], 10) || 0;
    const enteredBoxes = row.slice(4).filter(String).map(b => parseInt(b.toString().trim(), 10));
    
    vendorData.set(vendorCode, {
      vendorName: row[1],
      totalBoxes: totalBoxes,
      enteredBoxList: enteredBoxes,
      enteredCount: enteredBoxes.length
    });

    enteredBoxes.forEach(box => {
      if (!isNaN(box)) allEnteredBoxSet.add(box);
    });
  }
  return { vendorData, allEnteredBoxSet };
}

/**
 * 2. 「検品シート」からデータを取得し、条件に応じてフィルタリングする
 */
function fetchAndFilterInspectedData(today) {
  const data = INSPECTION_SHEET.getDataRange().getDisplayValues();
  const headers = data.shift();

  const colIndex = name => headers.indexOf(name);
  const colBoxNum = colIndex('箱番');
  const colTimestamp = colIndex('タイムスタンプ');
  const colIsMissing = colIndex('欠番');
  const colIsChecked = colIndex('検品済み');
  const colVendorCode = colIndex('業者名'); // 「業者コード」から「業者名」に変更
  const colAccount = colIndex('アカウント');
  const colScore = colIndex('点数');

  const baseFiltered = data.filter(row => 
    String(row[colIsMissing]).toUpperCase() !== 'TRUE' && 
    String(row[colIsChecked]).toUpperCase() === 'TRUE' && 
    row[colTimestamp]
  );
  
  const todayString = Utilities.formatDate(today, SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd");

  const todaysInspections = []; 
  const todaysInspectedBoxSet = new Set();

  baseFiltered.forEach(row => {
    const timestampString = row[colTimestamp];
    
    if (timestampString && timestampString.startsWith(todayString)) {
      const boxNum = parseInt(row[colBoxNum], 10);
      if(!isNaN(boxNum)) {
        todaysInspectedBoxSet.add(boxNum);
        todaysInspections.push({
          boxNum: boxNum,
          account: row[colAccount] || '（不明）',
          score: parseInt(row[colScore], 10) || 0
        });
      }
    }
  });

  const allInspectedBase = INSPECTION_SHEET.getDataRange().getValues();
  allInspectedBase.shift();
  const allInspectedData = allInspectedBase
    .filter(row => row[colIsMissing] !== true && row[colIsChecked] === true && row[colTimestamp])
    .map(row => ({
        boxNum: parseInt(row[colBoxNum], 10),
        vendorCode: row[colVendorCode] // ここで業者名がセットされる
    }));

  return { 
    todaysInspectedBoxSet: Array.from(todaysInspectedBoxSet),
    todaysInspections: todaysInspections,
    allInspectedData: allInspectedData
  };
}

/**
 * 3. 前日の結果シートから未入力リストを取得し、本日入力済みになったものを返す
 */
function checkRecoveredBoxes(allEnteredBoxSet) {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  const year = yesterday.getFullYear();
  const month = ('0' + (yesterday.getMonth() + 1)).slice(-2);
  const day = ('0' + yesterday.getDate()).slice(-2);

  const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const fileName = `${year}_${month}_検品チェック結果`;
  const files = folder.getFilesByName(fileName);
  
  if (!files.hasNext()) return [];

  const spreadsheet = SpreadsheetApp.open(files.next());
  const yesterdaySheet = spreadsheet.getSheetByName(`${day}日`);
  
  if (!yesterdaySheet) return [];

  const sheetData = yesterdaySheet.getDataRange().getValues();
  const previousUnentered = new Set();
  let headerIndex = -1;

  for(let i=0; i<sheetData.length; i++){
    const colIdx = sheetData[i].indexOf('本日検品→未入力の箱');
    if(colIdx > -1) {
      headerIndex = colIdx;
      for (let j = i + 1; j < sheetData.length; j++) {
        if (String(sheetData[j][0]).startsWith('■')) break; // 次のセクション
        const boxNum = parseInt(sheetData[j][headerIndex], 10);
        if(!isNaN(boxNum)) previousUnentered.add(boxNum);
      }
      break;
    }
  }

  const recovered = [];
  previousUnentered.forEach(boxNum => {
    if (allEnteredBoxSet.has(boxNum)) recovered.push(boxNum);
  });
  return recovered;
}

/**
 * 箱番シートを読み込み、「業者名」から「業者コード」への変換マップを作成する
 * @returns {Map<string, any>} 業者名 → 業者コード の変換マップ
 */
function createVendorCodeMapping() {
  if (!BOX_NUMBER_SHEET) {
    Logger.log("警告: '箱番シート'が見つかりません。業者コードの変換ができません。");
    return new Map();
  }
  const data = BOX_NUMBER_SHEET.getDataRange().getValues();
  const mapping = new Map();
  // ヘッダー行(1行目)をスキップ
  for (let i = 1; i < data.length; i++) {
    const vendorName = data[i][1]; // B列が業者名
    const vendorCode = data[i][2]; // C列が業者コード
    // まだ登録されていなければ、名前とコードのペアを登録
    if (vendorName && vendorCode && !mapping.has(vendorName)) {
      mapping.set(vendorName, vendorCode);
    }
  }
  return mapping;
}

/**
 * 4. 業者ごとの進捗を集計する (箱番シート経由で業者コードに変換する新ロジック)
 */
function calculateVendorProgress(enteredDataSummary, allInspectedData) {
  // 手順1: 検品シートから「業者名」ごとの検品数を集計
  const inspectedCountByName = new Map();
  allInspectedData.forEach(item => {
      const name = item.vendorCode; // この変数には業者名が格納されている
      const count = inspectedCountByName.get(name) || 0;
      inspectedCountByName.set(name, count + 1);
  });

  // 手順2: 箱番シートを使い、「業者名」→「業者コード」の変換マップを作成
  const vendorNameToCodeMap = createVendorCodeMapping();

  // 手順3: 「業者コード」をキーにした新しい検品数マップを作成
  const inspectedCountByCode = new Map();
  inspectedCountByName.forEach((count, name) => {
    const code = vendorNameToCodeMap.get(name);
    if (code) {
      inspectedCountByCode.set(code, count);
    } else {
      Logger.log(`警告: 箱番シートで業者名「${name}」に対応する業者コードが見つかりません。`);
    }
  });

  // 手順4: 入力済みマスターのデータと、「業者コード」を使って突合
  const progress = [];
  enteredDataSummary.vendorData.forEach((data, vendorCode) => {
    const inspectedCount = inspectedCountByCode.get(vendorCode) || 0;
    progress.push({
      code: vendorCode,
      name: data.vendorName,
      total: data.totalBoxes,
      entered: data.enteredCount,
      inspected: inspectedCount,
      enteredRatio: data.totalBoxes > 0 ? (data.enteredCount / data.totalBoxes * 100).toFixed(1) : 0,
      inspectedRatio: data.totalBoxes > 0 ? (inspectedCount / data.totalBoxes * 100).toFixed(1) : 0
    });
  });
  return progress;
}


/**
 * 5. 結果をスプレッドシートに書き出す (最終版)
 */
function generateReport(today, inspected, entered, unentered, recovered, progress, accountProgress, totalScoreToday) {
  const year = today.getFullYear();
  const month = ('0' + (today.getMonth() + 1)).slice(-2);
  const day = ('0' + today.getDate()).slice(-2);

  const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const fileName = `${year}_${month}_検品チェック結果`;
  
  let spreadsheet;
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    spreadsheet = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(spreadsheet.getId()).moveTo(folder);
  }

  const sheetName = `${day}日`;
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) spreadsheet.deleteSheet(sheet);
  sheet = spreadsheet.insertSheet(sheetName, 0);
  
  sheet.getRange("A1").setValue(`${year}/${month}/${day} 検品進捗レポート`).setFontSize(14).setFontWeight('bold');
  const dataStartRow = 3;

  // --- セクション1: 本日の状況 (A列から) ---
  sheet.getRange(dataStartRow, 1).setValue("■ 本日の状況").setFontSize(12).setFontWeight('bold').setBackground("#d9ead3");
  sheet.getRange(dataStartRow + 1, 1, 1, 3).setValues([['本日検品された箱', '現在入力済みの箱の総数', '本日検品→未入力の箱']]).setFontWeight('bold');
  
  const section1Data = [];
  const maxRows1 = Math.max(inspected.todaysInspectedBoxSet.length, entered.allEnteredBoxSet.size, unentered.length);
  for(let i = 0; i < maxRows1; i++) {
    section1Data.push([
      inspected.todaysInspectedBoxSet[i] || '',
      Array.from(entered.allEnteredBoxSet)[i] || '',
      unentered[i] || ''
    ]);
  }
  if (section1Data.length > 0) {
    sheet.getRange(dataStartRow + 2, 1, section1Data.length, 3).setValues(section1Data);
  }

  // --- セクション2: 前日からのリカバリー状況 (E列から) ---
  sheet.getRange(dataStartRow, 5).setValue("■ 前日未入力→本日入力済みの箱").setFontSize(12).setFontWeight('bold').setBackground("#fff2cc");
  if (recovered.length > 0) {
    sheet.getRange(dataStartRow + 1, 5, recovered.length, 1).setValues(recovered.map(r => [r]));
  } else {
    sheet.getRange(dataStartRow + 1, 5).setValue("該当なし");
  }

  // --- セクション3: 本日の作業進捗 (G列から) ---
  sheet.getRange(dataStartRow, 7).setValue("■ 本日の作業進捗").setFontSize(12).setFontWeight('bold').setBackground("#d9d2e9");
  sheet.getRange(dataStartRow + 1, 7, 2, 1).setValues([
    [`本日検品総数: ${inspected.todaysInspections.length} 箱`],
    [`本日検品総点数: ${totalScoreToday} 点`]
  ]).setFontWeight('bold');
  
  sheet.getRange(dataStartRow + 4, 7, 1, 3).setValues([['アカウント', '検品箱数', '合計点数']]).setFontWeight('bold');
  const accountProgressData = [];
  accountProgress.forEach((data, account) => {
    accountProgressData.push([account, data.boxCount, data.totalScore]);
  });
  if (accountProgressData.length > 0) {
    sheet.getRange(dataStartRow + 5, 7, accountProgressData.length, 3).setValues(accountProgressData);
  }
  
  // --- セクション4: 業者ごとの進捗サマリー (K列から) ---
  sheet.getRange(dataStartRow, 11).setValue("■ 業者ごと進捗サマリー").setFontSize(12).setFontWeight('bold').setBackground("#cfe2f3");
  const progressHeaders = ["コード", "業者名", "総箱数", "入力済み", "入力率(%)", "検品済み", "検品率(%)"];
  sheet.getRange(dataStartRow + 1, 11, 1, progressHeaders.length).setValues([progressHeaders]).setFontWeight('bold');
  
  const progressData = progress.map(p => [p.code, p.name, p.total, p.entered, p.enteredRatio, p.inspected, p.inspectedRatio]);
  if (progressData.length > 0) {
    sheet.getRange(dataStartRow + 2, 11, progressData.length, progressHeaders.length).setValues(progressData);
  }

  // 全ての列幅を自動調整
  if (sheet.getLastColumn() > 0) {
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  }
}
// '箱番シート'をグローバル変数として定義
const BOX_NUMBER_SHEET = SPREADSHEET.getSheetByName('箱番シート');

/**
 * ヘッダー名を基に列のインデックス（1始まり）を複数検索する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のシート
 * @param {string[]} columnNames - 検索する列名の配列
 * @returns {number[]} - 見つかった列のインデックスの配列
 */
function findColumnIndicesByName_(sheet, columnNames) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = [];
  columnNames.forEach(name => {
    const index = headers.indexOf(name);
    if (index !== -1) {
      indices.push(index + 1); // 1始まりのインデックス
    }
  });
  return indices;
}

/**
 * 毎月1日にトリガーで実行する関数。
 * 先月分のデータをアーカイブし、スプレッドシートのデータをクリアする。
 */
function archiveAndClearMonthlyData() {
  try {
    // 1. 先月の年月を計算する
    const now = new Date();
    now.setMonth(now.getMonth() - 1); // 月を1つ前に設定
    const archiveYearMonth = Utilities.formatDate(now, "Asia/Tokyo", "yyyy_MM");

    // 2. アーカイブ用のファイル名とフォルダパスを準備
    const archiveFolderName = `/検品チェック/画像/${archiveYearMonth}`;
    const archiveFileName = `${archiveYearMonth}_検品管理データベース`;

    // 3. アーカイブ先フォルダを取得または作成
    const destinationFolder = getOrCreateFolder(archiveFolderName);
    
    // 4. 現在のスプレッドシートをコピーしてアーカイブを作成
    const originalFile = DriveApp.getFileById(DATABASE_SHEET_ID);
    originalFile.makeCopy(archiveFileName, destinationFolder);
    Logger.log(`スプレッドシートのアーカイブを作成しました: ${archiveFolderName}/${archiveFileName}`);

    // 5. '検品シート'のデータをクリア
    if (INSPECTION_SHEET) {
      const inspectionColumnsToClear = ["検品済み", "欠番", "タイムスタンプ", "点数", "画像リンク", "アカウント"];
      const inspectionColumnIndices = findColumnIndicesByName_(INSPECTION_SHEET, inspectionColumnsToClear);
      
      if (inspectionColumnIndices.length > 0 && INSPECTION_SHEET.getLastRow() > 1) {
        inspectionColumnIndices.forEach(colIndex => {
          INSPECTION_SHEET.getRange(2, colIndex, INSPECTION_SHEET.getLastRow() - 1, 1).clearContent();
        });
        Logger.log('検品シートの指定列をクリアしました。');
      }
    } else {
      Logger.log('警告: 検品シートが見つかりません。');
    }

    // 6. '箱番シート'のデータをクリア
    if (BOX_NUMBER_SHEET) {
      const boxNumberColumnsToClear = ["箱番", "業者名", "業者コード"];
      const boxNumberColumnIndices = findColumnIndicesByName_(BOX_NUMBER_SHEET, boxNumberColumnsToClear);

      if (boxNumberColumnIndices.length > 0 && BOX_NUMBER_SHEET.getLastRow() > 1) {
        boxNumberColumnIndices.forEach(colIndex => {
          BOX_NUMBER_SHEET.getRange(2, colIndex, BOX_NUMBER_SHEET.getLastRow() - 1, 1).clearContent();
        });
        Logger.log('箱番シートの指定列をクリアしました。');
      }
    } else {
      Logger.log('警告: 箱番シートが見つかりません。');
    }

  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message}`);
  }
}
/**
 * fetchAndFilterInspectedData関数をテストするためだけの関数
 */
function test_myFunction() {
  // 1. 本来の処理と同様に、日付オブジェクトを作成する
  const today = new Date();
  
  // 2. 作成した日付を引数として渡し、関数を実行する
  const result = fetchAndFilterInspectedData(today);
  
  // 3. 結果をログに出力して確認する
  Logger.log(JSON.stringify(result, null, 2));
}
/**
 * 業者ごとの進捗計算が正しく行われるかテストする関数
 */
function test_VendorProgressCalculation() {
  try {
    // 手順1: マスターシートからサマリーを取得
    const enteredDataSummary = aggregateEnteredData();
    if (!enteredDataSummary) {
      Logger.log('入力済みマスターシートのデータ取得に失敗しました。');
      return;
    }

    // 手順2: 検品シートからデータを取得
    const inspectedDataSummary = fetchAndFilterInspectedData(new Date());
    if (!inspectedDataSummary) {
      Logger.log('検品シートのデータ取得に失敗しました。');
      return;
    }
    
    // 手順3: 目的の関数を実行して、業者ごとの進捗サマリーを計算
    const vendorProgress = calculateVendorProgress(enteredDataSummary, inspectedDataSummary.allInspectedData);

    // 手順4: 計算結果をログに出力
    Logger.log('--- 業者ごとの進捗サマリー 計算結果 ---');
    Logger.log(JSON.stringify(vendorProgress, null, 2));

  } catch(e) {
    Logger.log('テスト実行中にエラーが発生しました: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
  }
}