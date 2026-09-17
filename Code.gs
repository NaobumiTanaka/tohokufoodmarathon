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

const VALID_CATEGORIES = ['signage', 'caution', 'aid', 'guard', 'bus'];
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

// ---------------------------------------------------------------------------
// 初期データ一括登録（Googleマイマップ「エリア別_パイロン＆バーポイント／立哨ポイントMAP」からのインポート）
// ---------------------------------------------------------------------------
// 一度だけ実行する想定の関数。points シートの既存データはそのまま残し、末尾に追記する。
// 同じデータを二重登録しないよう、実行前に points シートを確認すること。
function importInitialPoints() {
  const sheet = getSheet_(POINTS_SHEET_NAME);
  const data = [{"area":"エリア1","label":"エリア1-1","memo":"","category":"caution","lat":38.695364,"lng":141.1428032},{"area":"エリア1","label":"エリア1-2","memo":"","category":"caution","lat":38.6987609,"lng":141.1454862},{"area":"エリア1","label":"エリア1-3","memo":"","category":"caution","lat":38.6987512,"lng":141.1454161},{"area":"エリア1","label":"エイド①","memo":"[エリア1]","category":"aid","lat":38.6990316,"lng":141.1456184},{"area":"エリア1","label":"BUS①","memo":"[エリア1]","category":"bus","lat":38.698571,"lng":141.1469381},{"area":"エリア1","label":"立哨2　（不要かも）","memo":"[エリア1] 不要かも","category":"guard","lat":38.7004601,"lng":141.1464886},{"area":"エリア1","label":"エリア1-4　（不要かも）","memo":"不要かも","category":"caution","lat":38.7003653,"lng":141.1469017},{"area":"エリア1","label":"立哨3　（不要かも）","memo":"[エリア1] 不要かも","category":"guard","lat":38.7003219,"lng":141.1469446},{"area":"エリア1","label":"立哨4　（不要かも）","memo":"[エリア1] 不要かも","category":"guard","lat":38.6939047,"lng":141.1448011},{"area":"エリア1","label":"エリア1-5","memo":"","category":"caution","lat":38.6939268,"lng":141.1448827},{"area":"エリア２","label":"エリア2-1","memo":"ラブホテル出入口　1／2","category":"caution","lat":38.6954241,"lng":141.1480876},{"area":"エリア２","label":"立哨5　（不要かも）","memo":"[エリア２] 不要かも","category":"guard","lat":38.6955459,"lng":141.1483416},{"area":"エリア２","label":"エリア2-2","memo":"ラブホテル出入口　２／2","category":"caution","lat":38.6956963,"lng":141.1485489},{"area":"エリア２","label":"エイド②","memo":"[エリア２]","category":"aid","lat":38.6965143,"lng":141.1518393},{"area":"エリア２","label":"立哨6","memo":"[エリア２]","category":"guard","lat":38.696922,"lng":141.1519309},{"area":"エリア２","label":"エリア2-3","memo":"","category":"caution","lat":38.696938,"lng":141.1518587},{"area":"エリア２","label":"BUS②","memo":"[エリア２]","category":"bus","lat":38.6971842,"lng":141.1537812},{"area":"エリア２","label":"立哨7","memo":"[エリア２]","category":"guard","lat":38.7006425,"lng":141.1524326},{"area":"エリア２","label":"エリア2-4","memo":"","category":"caution","lat":38.7006837,"lng":141.1522808},{"area":"エリア２","label":"立哨8　（不要かも）","memo":"[エリア２] 不要かも","category":"guard","lat":38.7009439,"lng":141.1516816},{"area":"エリア３","label":"立哨9","memo":"[エリア３]","category":"guard","lat":38.7006676,"lng":141.1549754},{"area":"エリア３","label":"エリア3-1","memo":"","category":"caution","lat":38.7005665,"lng":141.1550917},{"area":"エリア３","label":"エイド③","memo":"[エリア３]","category":"aid","lat":38.7049039,"lng":141.1538085},{"area":"エリア３","label":"BUS③","memo":"[エリア３]","category":"bus","lat":38.7052702,"lng":141.1563985},{"area":"エリア３","label":"エリア3-2","memo":"","category":"caution","lat":38.7052077,"lng":141.1566783},{"area":"エリア4","label":"立哨10","memo":"[エリア4]","category":"guard","lat":38.711618,"lng":141.1457381},{"area":"エリア4","label":"立哨11","memo":"[エリア4]","category":"guard","lat":38.7108553,"lng":141.143643},{"area":"エリア4","label":"エリア4-1","memo":"","category":"caution","lat":38.7116598,"lng":141.1457625},{"area":"エリア4","label":"エリア4-2","memo":"","category":"caution","lat":38.7108769,"lng":141.1436381},{"area":"エリア4","label":"エイド④","memo":"[エリア4]","category":"aid","lat":38.7093853,"lng":141.1413885},{"area":"エリア4","label":"エリア4-3","memo":"","category":"caution","lat":38.7093738,"lng":141.1415924},{"area":"エリア4","label":"エリア4-4","memo":"","category":"caution","lat":38.7092189,"lng":141.1410264},{"area":"エリア4","label":"エリア4-5","memo":"","category":"caution","lat":38.7089019,"lng":141.1399523},{"area":"エリア4","label":"立哨12","memo":"[エリア4]","category":"guard","lat":38.7064594,"lng":141.1390589},{"area":"エリア4","label":"BUS4","memo":"[エリア4]","category":"bus","lat":38.7088226,"lng":141.1396825},{"area":"エリア5","label":"エリア5-1","memo":"","category":"caution","lat":38.7043268,"lng":141.1369298},{"area":"エリア5","label":"立哨13","memo":"[エリア5]","category":"guard","lat":38.7043739,"lng":141.1367563},{"area":"エリア5","label":"エリア5-2","memo":"","category":"caution","lat":38.7028366,"lng":141.1354829},{"area":"エリア5","label":"立哨14","memo":"[エリア5]","category":"guard","lat":38.7020587,"lng":141.1336885},{"area":"エリア5","label":"エリア5-3","memo":"","category":"caution","lat":38.7010136,"lng":141.1317949},{"area":"エリア5","label":"立哨15","memo":"[エリア5]","category":"guard","lat":38.7009969,"lng":141.1317747},{"area":"エリア5","label":"BUS5","memo":"[エリア5]","category":"bus","lat":38.7007681,"lng":141.1300764},{"area":"エリア5","label":"エイド⑤","memo":"[エリア5]","category":"aid","lat":38.7010329,"lng":141.1323057},{"area":"エリア5","label":"エリア5-4","memo":"","category":"caution","lat":38.6972107,"lng":141.1266311},{"area":"エリア5","label":"エリア5-5","memo":"","category":"caution","lat":38.6970359,"lng":141.1266163},{"area":"エリア5","label":"立哨16","memo":"[エリア5]","category":"guard","lat":38.6972526,"lng":141.1266807},{"area":"エリア5","label":"ポイント 12","memo":"","category":"caution","lat":38.7054702,"lng":141.1368869},{"area":"エリア6","label":"立哨17","memo":"[エリア6]","category":"guard","lat":38.6952687,"lng":141.1238464},{"area":"エリア6","label":"立哨18","memo":"[エリア6]","category":"guard","lat":38.6942103,"lng":141.1210126},{"area":"エリア6","label":"エリア6-1","memo":"","category":"caution","lat":38.6953273,"lng":141.1237579},{"area":"エリア6","label":"エリア6-2","memo":"","category":"caution","lat":38.6942626,"lng":141.1209536},{"area":"エリア6","label":"エイド⑥","memo":"[エリア6]","category":"aid","lat":38.6932452,"lng":141.1205325},{"area":"エリア6","label":"BUS6","memo":"[エリア6]","category":"bus","lat":38.695574,"lng":141.1206861},{"area":"エリア6","label":"エリア6-3","memo":"","category":"caution","lat":38.6927929,"lng":141.1162003},{"area":"エリア6","label":"立哨19","memo":"[エリア6]","category":"guard","lat":38.6928075,"lng":141.1162486},{"area":"エリア7","label":"エイド⑦","memo":"[エリア7] GoogleMapと実際の道路が違う","category":"aid","lat":38.6906385,"lng":141.1064702},{"area":"エリア7","label":"BUS7","memo":"[エリア7]","category":"bus","lat":38.694259,"lng":141.1066542},{"area":"エリア7","label":"エリア7-1（設置しなくてもよい）","memo":"ここは設置しなくてもよい","category":"caution","lat":38.6925351,"lng":141.1094794},{"area":"エリア7","label":"エリア7-2","memo":"北方向にパイロン設置。西方向には道路は無い","category":"caution","lat":38.6906332,"lng":141.1063271},{"area":"エリア7","label":"エリア7-3","memo":"","category":"caution","lat":38.6923469,"lng":141.1054515},{"area":"エリア7","label":"立哨21（設置しなくてもよい）","memo":"[エリア7] ここは設置しなくてもよい","category":"guard","lat":38.6925251,"lng":141.1094511},{"area":"エリア7","label":"立哨22","memo":"[エリア7]","category":"guard","lat":38.6906151,"lng":141.106406},{"area":"エリア7","label":"立哨23","memo":"[エリア7] 釣り客が道路横断するので注意","category":"guard","lat":38.6879745,"lng":141.1063401},{"area":"エリア7","label":"立哨24（不要かも）","memo":"[エリア7] ここは不要かも","category":"guard","lat":38.6871669,"lng":141.1068125},{"area":"エリア8","label":"エリア8-1a","memo":"ポイント 27-2","category":"caution","lat":38.6859886,"lng":141.1101851},{"area":"エリア8","label":"エリア8-1b","memo":"ポイント 27-1","category":"caution","lat":38.6860534,"lng":141.1102681},{"area":"エリア8","label":"エリア8-2a","memo":"ポイント28-1","category":"caution","lat":38.6836794,"lng":141.1134805},{"area":"エリア8","label":"エリア8-2b","memo":"ポイント 28-2","category":"caution","lat":38.6835748,"lng":141.1135411},{"area":"エリア8","label":"エリア8-3","memo":"ポイント 29 / 注意：北東方面にいく道路は閉鎖中","category":"caution","lat":38.6816761,"lng":141.1159715},{"area":"エリア8","label":"立哨25","memo":"[エリア8] 交差点のため、自動車さばき注意","category":"guard","lat":38.6860837,"lng":141.1102251},{"area":"エリア8","label":"立哨26（不要かも）","memo":"[エリア8] ここは要らないかも","category":"guard","lat":38.6881173,"lng":141.1117308},{"area":"エリア8","label":"立哨27","memo":"[エリア8] 立哨26をなくす場合は、もう少し、エイド⑦寄りに設置する","category":"guard","lat":38.6872591,"lng":141.1183864},{"area":"エリア8","label":"立哨28","memo":"[エリア8]","category":"guard","lat":38.6849062,"lng":141.1157667},{"area":"エリア8","label":"立哨29","memo":"[エリア8]","category":"guard","lat":38.6817085,"lng":141.1159213},{"area":"エリア8","label":"エイド⑧","memo":"[エリア8]","category":"aid","lat":38.6836578,"lng":141.1136676},{"area":"エリア8","label":"BUS⑧","memo":"[エリア8]","category":"bus","lat":38.6835888,"lng":141.1134669},{"area":"エリア9","label":"立哨30","memo":"[エリア9] 交差点のため、2方向からの自動車さばきに注意","category":"guard","lat":38.6796995,"lng":141.1167406},{"area":"エリア9","label":"立哨31","memo":"[エリア9]","category":"guard","lat":38.6789703,"lng":141.119469},{"area":"エリア9","label":"立哨（新規設置候補）","memo":"[エリア9] 北方向に走ってしまう人数名あり。立哨の新規設置検討","category":"guard","lat":38.6819245,"lng":141.1230958},{"area":"エリア9","label":"立哨32","memo":"[エリア9]","category":"guard","lat":38.6797532,"lng":141.1237757},{"area":"エリア9","label":"立哨33","memo":"[エリア9]","category":"guard","lat":38.6794904,"lng":141.1224346},{"area":"エリア9","label":"立哨34","memo":"[エリア9]","category":"guard","lat":38.6814814,"lng":141.1298075},{"area":"エリア9","label":"エイド⑨","memo":"[エリア9]","category":"aid","lat":38.6798149,"lng":141.1237019},{"area":"エリア9","label":"エリア9-1a","memo":"","category":"caution","lat":38.6796744,"lng":141.116785},{"area":"エリア9","label":"エリア9-1b","memo":"","category":"caution","lat":38.6797163,"lng":141.1166977},{"area":"エリア9","label":"エリア9-2a","memo":"","category":"caution","lat":38.67897,"lng":141.1194146},{"area":"エリア9","label":"エリア9-2b","memo":"ガードレール設置済みのため、パイロン不要","category":"caution","lat":38.6791187,"lng":141.1193865},{"area":"エリア9","label":"エリア9-3a","memo":"","category":"caution","lat":38.6802133,"lng":141.1218892},{"area":"エリア9","label":"エリア9-3b","memo":"","category":"caution","lat":38.6802803,"lng":141.1217416},{"area":"エリア9","label":"エリア9-4","memo":"","category":"caution","lat":38.6815823,"lng":141.1223169},{"area":"エリア9","label":"エリア9-5a","memo":"","category":"caution","lat":38.6819622,"lng":141.1230703},{"area":"エリア9","label":"エリア9-5b","memo":"","category":"caution","lat":38.6818292,"lng":141.1229964},{"area":"エリア9","label":"エリア9-6","memo":"","category":"caution","lat":38.6796726,"lng":141.1237891},{"area":"エリア9","label":"エリア9-7","memo":"","category":"caution","lat":38.6813909,"lng":141.1298765},{"area":"エリア9","label":"BUS⑨-1","memo":"[エリア9]","category":"bus","lat":38.6783696,"lng":141.1213627},{"area":"エリア9","label":"BUS⑨-2","memo":"[エリア9]","category":"bus","lat":38.6814022,"lng":141.1299939},{"area":"エリア10","label":"立哨35","memo":"[エリア10] 駐車スペースあり、釣り客注意","category":"guard","lat":38.6845492,"lng":141.1278369},{"area":"エリア10","label":"立哨36","memo":"[エリア10]","category":"guard","lat":38.686793,"lng":141.1297769},{"area":"エリア10","label":"立哨37","memo":"[エリア10]","category":"guard","lat":38.6902688,"lng":141.1317085},{"area":"エリア10","label":"立哨38","memo":"[エリア10]","category":"guard","lat":38.6868151,"lng":141.1318759},{"area":"エリア10","label":"エイド⑩","memo":"[エリア10]","category":"aid","lat":38.6878596,"lng":141.1312106},{"area":"エリア10","label":"エリア10-1","memo":"","category":"caution","lat":38.6867386,"lng":141.1297474},{"area":"エリア10","label":"新規パイロン設置候補エリア10-1.5","memo":"余裕あればパイロン設置","category":"caution","lat":38.6906891,"lng":141.1300316},{"area":"エリア10","label":"エリア10-2a","memo":"","category":"caution","lat":38.6878177,"lng":141.1311355},{"area":"エリア10","label":"エリア10-2b","memo":"","category":"caution","lat":38.6878009,"lng":141.1313126},{"area":"エリア10","label":"エリア10-3","memo":"","category":"caution","lat":38.6868037,"lng":141.1317922},{"area":"エリア10","label":"BUS⑩","memo":"[エリア10]","category":"bus","lat":38.6890073,"lng":141.1351691}];

  const now = new Date();
  const rows = data.map(function (d, i) {
    const id = 'pt_' + now.getTime() + '_' + i + '_' + Math.floor(Math.random() * 100000);
    return [id, d.category, d.lat, d.lng, d.label, 'not_placed', '', d.memo, now, now];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, POINTS_HEADERS.length).setValues(rows);

  Logger.log('初期データ登録完了: ' + rows.length + '件');
}
