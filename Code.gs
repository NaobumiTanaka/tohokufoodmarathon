/**
 * 東北風土マラソン - コース看板・矢印 設置管理アプリ バックエンド (Google Apps Script)
 *
 * harunojin (越後まつだい春の陣) の arrows.html / mow.html と同じアーキテクチャを踏襲。
 * - スプレッドシートを簡易DBとして使用
 * - Web App として公開し、フロントエンド(signage.html)から JSONP で読み書きする
 *
 * ===== セットアップ手順 =====
 * 1. Google スプレッドシートを新規作成する
 * 2. 拡張機能 → Apps Script を開き、このファイルの内容を貼り付ける
 * 3. 下の SHEET_ID に、作成したスプレッドシートのIDを設定する
 *    (スプレッドシートのURL: https://docs.google.com/spreadsheets/d/【ここがID】/edit)
 * 4. エディタ上部の関数選択で initializeSheets を選び、一度実行する
 *    (初回は権限承認が必要)
 * 5. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 *    でデプロイし、発行されたURLを signage.html の API_URL に設定する
 *
 * ===== 再デプロイ時の注意 =====
 * コードだけ更新する場合は「デプロイを管理」→ 鉛筆アイコン →
 * バージョン「新バージョン」→ デプロイ（URLは変わらない）。
 * 「新しいデプロイ」を選ぶと新しいURLが発行され、signage.html側のAPI_URLも
 * 更新が必要になるので注意。
 */

const SHEET_ID = 'ここにスプレッドシートIDを入力'; // TODO: 実際のスプレッドシートIDに置き換える
const POINTS_SHEET_NAME = 'points';
const LOG_SHEET_NAME = 'log';
const COURSE_SHEET_NAME = 'course';

const POINTS_HEADERS = [
  'id', 'category', 'lat', 'lng', 'label', 'status', 'worker', 'memo',
  'created_at', 'updated_at'
];
const LOG_HEADERS = [
  'timestamp', 'action', 'id', 'category', 'status', 'worker', 'memo'
];
const COURSE_HEADERS = ['seq', 'lat', 'lng', 'updated_at'];

const VALID_CATEGORIES = ['signage', 'caution', 'aid'];
const VALID_STATUSES = ['not_placed', 'placed', 'collected'];

// ---------------------------------------------------------------------------
// エントリーポイント
// ---------------------------------------------------------------------------

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action || 'get';
  let result;

  try {
    if (action === 'get') {
      result = { success: true, points: getPoints() };
    } else if (action === 'add') {
      result = addPoint(params);
    } else if (action === 'update') {
      result = updatePoint(params);
    } else if (action === 'delete') {
      result = deletePoint(params);
    } else if (action === 'getCourse') {
      result = { success: true, course: getCourse() };
    } else if (action === 'clearCourse') {
      result = clearCourse();
    } else if (action === 'addCourseBatch') {
      result = addCourseBatch(params);
    } else {
      result = { success: false, error: '不明なaction: ' + action };
    }
  } catch (err) {
    result = { success: false, error: String(err && err.message ? err.message : err) };
  }

  return respond(result, params.callback);
}

function respond(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// 初期化
// ---------------------------------------------------------------------------

function initializeSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let pointsSheet = ss.getSheetByName(POINTS_SHEET_NAME);
  if (!pointsSheet) {
    pointsSheet = ss.insertSheet(POINTS_SHEET_NAME);
  }
  if (pointsSheet.getLastRow() === 0) {
    pointsSheet.appendRow(POINTS_HEADERS);
  }

  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
  }
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(LOG_HEADERS);
  }

  let courseSheet = ss.getSheetByName(COURSE_SHEET_NAME);
  if (!courseSheet) {
    courseSheet = ss.insertSheet(COURSE_SHEET_NAME);
  }
  if (courseSheet.getLastRow() === 0) {
    courseSheet.appendRow(COURSE_HEADERS);
  }

  // デフォルトで作成される「シート1」が残っていれば削除
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 3) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('初期化完了: points / log シートを用意しました');
}

// ---------------------------------------------------------------------------
// データ操作
// ---------------------------------------------------------------------------

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + name + ' (initializeSheetsを実行してください)');
  }
  return sheet;
}

function getPoints() {
  const sheet = getSheet_(POINTS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  const rows = values.slice(1);

  return rows
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function findRowIndexById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // 1-indexed row number
  }
  return -1;
}

function addPoint(params) {
  const sheet = getSheet_(POINTS_SHEET_NAME);

  const category = params.category;
  const lat = parseFloat(params.lat);
  const lng = parseFloat(params.lng);

  if (VALID_CATEGORIES.indexOf(category) === -1) {
    return { success: false, error: '不正なcategory: ' + category };
  }
  if (isNaN(lat) || isNaN(lng)) {
    return { success: false, error: '不正な座標です' };
  }

  const id = 'pt_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
  const now = new Date();
  const status = VALID_STATUSES.indexOf(params.status) !== -1 ? params.status : 'not_placed';

  sheet.appendRow([
    id,
    category,
    lat,
    lng,
    params.label || '',
    status,
    params.worker || '',
    params.memo || '',
    now,
    now
  ]);

  logAction_('add', id, category, status, params.worker || '', params.memo || '');

  return { success: true, id: id };
}

function updatePoint(params) {
  const sheet = getSheet_(POINTS_SHEET_NAME);
  const id = params.id;
  if (!id) {
    return { success: false, error: 'idが指定されていません' };
  }

  const rowIndex = findRowIndexById_(sheet, id);

  if (rowIndex === -1) {
    // upsert: 見つからない場合は新規追加にフォールバック
    if (params.lat !== undefined && params.lng !== undefined && params.category) {
      return addPoint(params);
    }
    return { success: false, error: 'id が見つかりません: ' + id };
  }

  const headers = POINTS_HEADERS;
  const colOf = function (name) { return headers.indexOf(name) + 1; };

  if (params.status !== undefined) {
    if (VALID_STATUSES.indexOf(params.status) === -1) {
      return { success: false, error: '不正なstatus: ' + params.status };
    }
    sheet.getRange(rowIndex, colOf('status')).setValue(params.status);
  }
  if (params.worker !== undefined) {
    sheet.getRange(rowIndex, colOf('worker')).setValue(params.worker);
  }
  if (params.memo !== undefined) {
    sheet.getRange(rowIndex, colOf('memo')).setValue(params.memo);
  }
  if (params.label !== undefined) {
    sheet.getRange(rowIndex, colOf('label')).setValue(params.label);
  }
  if (params.lat !== undefined) {
    sheet.getRange(rowIndex, colOf('lat')).setValue(parseFloat(params.lat));
  }
  if (params.lng !== undefined) {
    sheet.getRange(rowIndex, colOf('lng')).setValue(parseFloat(params.lng));
  }
  sheet.getRange(rowIndex, colOf('updated_at')).setValue(new Date());

  logAction_('update', id, params.category || '', params.status || '', params.worker || '', params.memo || '');

  return { success: true, id: id };
}

function deletePoint(params) {
  const sheet = getSheet_(POINTS_SHEET_NAME);
  const id = params.id;
  if (!id) {
    return { success: false, error: 'idが指定されていません' };
  }
  const rowIndex = findRowIndexById_(sheet, id);
  if (rowIndex === -1) {
    return { success: false, error: 'id が見つかりません: ' + id };
  }
  sheet.deleteRow(rowIndex);
  logAction_('delete', id, '', '', '', '');
  return { success: true, id: id };
}

function logAction_(action, id, category, status, worker, memo) {
  try {
    const sheet = getSheet_(LOG_SHEET_NAME);
    sheet.appendRow([new Date(), action, id, category, status, worker, memo]);
  } catch (err) {
    // ログ失敗はメイン処理に影響させない
    Logger.log('ログ書き込み失敗: ' + err);
  }
}

// ---------------------------------------------------------------------------
// コースライン（アプリの地図上で手描きして保存する）
// ---------------------------------------------------------------------------

function getCourse() {
  const sheet = getSheet_(COURSE_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const rows = values.slice(1)
    .filter(function (row) { return row[1] !== '' && row[1] !== null; })
    .map(function (row) {
      return { seq: Number(row[0]), lat: Number(row[1]), lng: Number(row[2]) };
    });

  rows.sort(function (a, b) { return a.seq - b.seq; });
  return rows;
}

function clearCourse() {
  const sheet = getSheet_(COURSE_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, COURSE_HEADERS.length).clearContent();
  }
  return { success: true };
}

// params.points は JSON文字列: [{"seq":0,"lat":38.6,"lng":141.1}, ...]
// クリック数が多いコースは複数バッチに分けて呼び出す想定(1回あたり数十点)。
// clearCourse を呼んでから追記する運用なので、この関数自体は常に「追記」のみ行う。
function addCourseBatch(params) {
  const sheet = getSheet_(COURSE_SHEET_NAME);
  let batch;
  try {
    batch = JSON.parse(params.points);
  } catch (err) {
    return { success: false, error: 'points の形式が不正です' };
  }
  if (!Array.isArray(batch) || batch.length === 0) {
    return { success: false, error: 'points が空です' };
  }

  const now = new Date();
  const rows = batch.map(function (p) {
    return [Number(p.seq), Number(p.lat), Number(p.lng), now];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, COURSE_HEADERS.length).setValues(rows);

  return { success: true, added: rows.length };
}
