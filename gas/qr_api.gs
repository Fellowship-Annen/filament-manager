function doPost(e) {

  try {
    const data = JSON.parse(e.postData.contents);
    const action = String(data.action || "").trim();

    if (action === "markPrinted") return markPrinted(data);
    if (action === "updateLabel") return updateLabel(data);
    if (action === "createFilamentMaster") return createFilamentMaster(data);

    // ==============================
    // 重量更新
    // ==============================
    if (action === "updateLocation") return updateLocation(data);
    if (action === "updateWeight") {
      return updateWeight(data);
    }

    // ==============================
    // 新規登録
    // ==============================
    if (action === "register") {
      return registerFilament(data);
    }

    // 旧テストページとの互換性
    // action無しで id + weight が来た場合は重量更新
    if (!action && data.id && data.weight !== undefined) {
      return updateWeight(data);
    }

    return jsonResponse({
      ok: false,
      error: "UNKNOWN_ACTION"
    });

  } catch (error) {

    return jsonResponse({
      ok: false,
      error: error.message
    });
  }
}

function updateWeight(data) { return mutateWithHistory_(data, 'updateWeight'); }
function updateLocation(data) { return mutateWithHistory_(data, 'updateLocation'); }
function registerFilament(data) { return mutateWithHistory_(data, 'register'); }

function doGet(e) {

  try {
    if (e && e.parameter && e.parameter.action === 'health') {
      return jsonResponse({ok: true, apiVersion: 'filament-form-history-v3', historySheet: 'フォームの回答 1', actions: ['updateWeight', 'updateLocation', 'register', 'markPrinted', 'updateLabel', 'filamentMasters', 'makerMasters', 'createFilamentMaster']});
    }
    if (e && e.parameter && e.parameter.action === 'filamentMasters') {
      return getFilamentMasters_();
    }
    if (e && e.parameter && e.parameter.action === 'makerMasters') {
      return getMakerMasters_();
    }
    const id = String(e.parameter.id || "").trim();

    if (!id) {
      return jsonResponse({
        ok: false,
        error: "ID_REQUIRED"
      });
    }

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("在庫台帳");

    if (!sheet) {
      throw new Error("在庫台帳が見つかりません");
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return jsonResponse({
        ok: false,
        error: "ID_NOT_FOUND"
      });
    }

    // A～I列をまとめて取得
    const data = sheet
      .getRange(2, 1, lastRow - 1, 9)
      .getValues();

    for (let i = 0; i < data.length; i++) {

      if (String(data[i][0]).trim() === id) {

        return jsonResponse({
          ok: true,

          id:       data[i][0], // A 管理番号
          maker:    data[i][1], // B メーカー
          base:     data[i][2], // C 母材
          sub:      data[i][3], // D サブカテゴリ
          color:    data[i][4], // E 色
          location: data[i][5], // F 保管場所
          weight:   data[i][6], // G 重量
          person:   data[i][7], // H 担当
          note:     data[i][8]  // I 備考
        });
      }
    }

    return jsonResponse({
      ok: false,
      error: "ID_NOT_FOUND"
    });

  } catch (error) {

    return jsonResponse({
      ok: false,
      error: error.message
    });
  }
}

function getFilamentMasters_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const makerSheet = book.getSheetByName('メーカーマスター');
  const filamentSheet = book.getSheetByName('フィラメントマスター');
  if (!makerSheet || !filamentSheet) throw new Error('メーカーマスターまたはフィラメントマスターが見つかりません');
  const makerRows = makerSheet.getDataRange().getValues();
  const filamentRows = filamentSheet.getDataRange().getValues();
  const makerHeaders = makerRows[0].map(v => historyText_(v).normalize('NFKC').replace(/\s/g, ''));
  const filamentHeaders = filamentRows[0].map(v => historyText_(v).normalize('NFKC').replace(/\s/g, ''));
  if (!/^メーカーID$/.test(makerHeaders[0]) || !/^メーカー名$/.test(makerHeaders[1]) || !/標準空スプール重量/.test(makerHeaders[2])) throw new Error('メーカーマスターの列を確認してください');
  const expected = [/^フィラメントID$/, /^メーカーID$/, /商品名|シリーズ名/, /^素材$/, /サブカテゴリ/, /^色$/, /検索別名/, /購入先URL/, /備考/];
  if (expected.some((test, index) => !test.test(filamentHeaders[index] || ''))) throw new Error('フィラメントマスターの列を確認してください');
  const makers = new Map();
  makerRows.slice(1).forEach(row => {
    const makerId = historyText_(row[0]).toUpperCase();
    if (makerId) makers.set(makerId, {name: historyText_(row[1]), spoolWeight: row[2] === '' ? null : Number(row[2])});
  });
  const items = filamentRows.slice(1).map(row => {
    const id = historyText_(row[0]).toUpperCase(), makerId = historyText_(row[1]).toUpperCase(), maker = makers.get(makerId);
    if (!id && !makerId) return null;
    if (!/^FL\d{6}$/.test(id) || !maker) throw new Error('フィラメントマスターのIDまたはメーカー参照を確認してください');
    return {id:id, makerId:makerId, maker:maker.name, product:historyText_(row[2]), base:historyText_(row[3]), sub:historyText_(row[4]), color:historyText_(row[5]), aliases:historyText_(row[6]), purchaseUrl:historyText_(row[7]), note:historyText_(row[8]), spoolWeight:Number.isFinite(maker.spoolWeight) ? maker.spoolWeight : null};
  }).filter(Boolean);
  return jsonResponse({ok:true, apiVersion:'filament-form-history-v3', action:'filamentMasters', items:items});
}

function masterHeaders_(sheet, tests, name) {
  if (!sheet) throw new Error(name + 'が見つかりません');
  const width = tests.length;
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0]
    .map(v => historyText_(v).normalize('NFKC').replace(/\s/g, ''));
  if (tests.some((test, index) => !test.test(headers[index] || ''))) {
    throw new Error(name + 'の列を確認してください');
  }
  return headers;
}

function getMakerMasters_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('メーカーマスター');
  masterHeaders_(sheet, [/^メーカーID$/, /^メーカー名$/, /標準空スプール重量/, /メーカー別名/], 'メーカーマスター');
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];
  const ids = new Set(), names = new Set();
  const items = rows.map(row => {
    const id = historyText_(row[0]).toUpperCase(), name = historyText_(row[1]);
    if (!id && !name) return null;
    if (!/^MK\d{3}$/.test(id) || !name) throw new Error('メーカーマスターのIDまたは名称を確認してください');
    const normalizedName = name.normalize('NFKC').toLowerCase();
    if (ids.has(id) || names.has(normalizedName)) throw new Error('メーカーマスターに重複があります');
    ids.add(id); names.add(normalizedName);
    const rawWeight = row[2], spoolWeight = rawWeight === '' ? null : Number(rawWeight);
    if (spoolWeight !== null && (!Number.isFinite(spoolWeight) || spoolWeight < 0 || spoolWeight > 2000)) throw new Error('標準空スプール重量を確認してください');
    return {id:id, name:name, spoolWeight:spoolWeight, aliases:historyText_(row[3])};
  }).filter(Boolean);
  return jsonResponse({ok:true, apiVersion:'filament-form-history-v3', action:'makerMasters', items:items});
}

function masterValue_(value, name, required, maxLength) {
  const text = historyText_(value);
  if (required && !text) throw new Error(name + 'を入力してください');
  if (text.length > maxLength || /[\r\n]/.test(text)) throw new Error(name + 'は改行せず' + maxLength + '文字以内で入力してください');
  return text;
}

function createFilamentMaster(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const makerSheet = book.getSheetByName('メーカーマスター');
    const filamentSheet = book.getSheetByName('フィラメントマスター');
    masterHeaders_(makerSheet, [/^メーカーID$/, /^メーカー名$/, /標準空スプール重量/, /メーカー別名/], 'メーカーマスター');
    masterHeaders_(filamentSheet, [/^フィラメントID$/, /^メーカーID$/, /商品名|シリーズ名/, /^素材$/, /サブカテゴリ/, /^色$/, /検索別名/, /購入先URL/, /備考/], 'フィラメントマスター');

    const makerId = masterValue_(data.makerId, 'メーカー', true, 5).toUpperCase();
    if (!/^MK\d{3}$/.test(makerId)) throw new Error('メーカーIDを確認してください');
    const makerRows = makerSheet.getLastRow() > 1 ? makerSheet.getRange(2, 1, makerSheet.getLastRow() - 1, 2).getValues() : [];
    const makerMatches = makerRows.filter(row => historyText_(row[0]).toUpperCase() === makerId);
    if (makerMatches.length !== 1) throw new Error(makerMatches.length ? 'メーカーIDが重複しています' : 'メーカーが見つかりません');

    const product = masterValue_(data.product, '商品名・シリーズ名', false, 100);
    const base = masterValue_(data.base, '素材', true, 50);
    const sub = masterValue_(data.sub, 'サブカテゴリ', true, 50);
    const color = masterValue_(data.color, '色', true, 80);
    const aliases = masterValue_(data.aliases, '検索別名', false, 300);
    const purchaseUrl = masterValue_(data.purchaseUrl, '購入先URL', false, 500);
    const note = masterValue_(data.note, '備考', false, 1000);
    if (purchaseUrl && !/^https:\/\//i.test(purchaseUrl)) throw new Error('購入先URLはhttps://から入力してください');

    const lastRow = filamentSheet.getLastRow();
    const rows = lastRow > 1 ? filamentSheet.getRange(2, 1, lastRow - 1, 9).getValues() : [];
    const normalized = value => historyText_(value).normalize('NFKC').toLowerCase();
    const duplicate = rows.find(row => historyText_(row[1]).toUpperCase() === makerId && normalized(row[3]) === normalized(base) && normalized(row[4]) === normalized(sub) && normalized(row[5]) === normalized(color));
    if (duplicate) {
      return jsonResponse({ok:false, error:'MASTER_ALREADY_EXISTS', message:'同じメーカー・素材・サブカテゴリ・色のフィラメントがすでにあります', existingId:historyText_(duplicate[0]).toUpperCase()});
    }
    const numbers = rows.map(row => historyText_(row[0]).toUpperCase()).filter(id => /^FL\d{6}$/.test(id)).map(id => Number(id.slice(2)));
    if (rows.some(row => historyText_(row[0]) && !/^FL\d{6}$/.test(historyText_(row[0]).toUpperCase()))) throw new Error('フィラメントマスターのID形式を確認してください');
    const next = Math.max(0, ...numbers) + 1;
    if (next > 999999) throw new Error('フィラメントIDの上限に達しました');
    const id = 'FL' + String(next).padStart(6, '0');
    if (rows.some(row => historyText_(row[0]).toUpperCase() === id)) throw new Error('採番したフィラメントIDが重複しています');
    filamentSheet.appendRow([id, makerId, product, base, sub, color, aliases, purchaseUrl, note].map(historyCell_));
    SpreadsheetApp.flush();
    return jsonResponse({ok:true, apiVersion:'filament-form-history-v3', action:'createFilamentMaster', id:id, makerId:makerId, maker:historyText_(makerMatches[0][1]), product:product, base:base, sub:sub, color:color, aliases:aliases, purchaseUrl:purchaseUrl, note:note});
  } catch (error) {
    return jsonResponse({ok:false, error:String(error.message), message:String(error.message)});
  } finally {
    lock.releaseLock();
  }
}

function requireFilamentMaster_(book, filamentKey) {
  const key = historyText_(filamentKey).toUpperCase();
  if (!/^FL\d{6}$/.test(key)) throw new Error('フィラメントを選択してください');
  const makerSheet = book.getSheetByName('メーカーマスター');
  const filamentSheet = book.getSheetByName('フィラメントマスター');
  if (!makerSheet || !filamentSheet) throw new Error('メーカーマスターまたはフィラメントマスターが見つかりません');
  const makers = new Map();
  const makerLastRow = makerSheet.getLastRow();
  if (makerLastRow > 1) makerSheet.getRange(2, 1, makerLastRow - 1, 2).getValues().forEach(row => makers.set(historyText_(row[0]).toUpperCase(), historyText_(row[1])));
  const lastRow = filamentSheet.getLastRow();
  if (lastRow < 2) throw new Error('フィラメントマスターが空です');
  const matches = filamentSheet.getRange(2, 1, lastRow - 1, 9).getValues().filter(row => historyText_(row[0]).toUpperCase() === key);
  if (matches.length !== 1) throw new Error(matches.length ? 'フィラメントIDが重複しています' : 'フィラメントIDが見つかりません');
  const row = matches[0], makerId = historyText_(row[1]).toUpperCase(), maker = makers.get(makerId);
  if (!maker) throw new Error('フィラメントのメーカーIDがメーカーマスターにありません');
  return {filamentKey:key, maker:maker, base:historyText_(row[3]), sub:historyText_(row[4]), color:historyText_(row[5])};
}

function jsonResponse(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 印刷用の短い注釈をL列へ保存する。在庫内容の変更ではないため、フォーム回答には追記しない。
function updateLabel(data) {
  const id = String(data.id || '').trim();
  const label = String(data.label == null ? '' : data.label).trim();
  if (!/^[SOF]\d{8}$/.test(id)) {
    return jsonResponse({ok: false, error: 'INVALID_ID', message: '管理番号を確認してください'});
  }
  if (label.length > 60 || /[\r\n]/.test(label)) {
    return jsonResponse({ok: false, error: 'INVALID_LABEL', message: 'ラベル注釈は改行せず60文字以内で入力してください'});
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫台帳');
    if (!ledger) throw new Error('在庫台帳が見つかりません');
    const header = String(ledger.getRange(1, 12).getValue() || '').trim().normalize('NFKC').replace(/\s/g, '');
    if (!/ラベル注釈/.test(header)) throw new Error('在庫台帳L列の見出し「ラベル注釈」を確認してください');
    const lastRow = ledger.getLastRow();
    if (lastRow < 2) throw new Error('ID_NOT_FOUND');
    const ids = ledger.getRange(2, 1, lastRow - 1, 1).getValues();
    const matches = ids.map((row, index) => String(row[0]).trim() === id ? index + 2 : -1).filter(row => row > 0);
    if (matches.length !== 1) throw new Error(matches.length ? 'DUPLICATE_ID' : 'ID_NOT_FOUND');
    ledger.getRange(matches[0], 12).setValue(historyCell_(label));
    SpreadsheetApp.flush();
    return jsonResponse({ok: true, apiVersion: 'filament-form-history-v3', action: 'updateLabel', id: id, label: label});
  } catch (error) {
    return jsonResponse({ok: false, error: String(error.message), message: 'ラベル注釈を保存できませんでした'});
  } finally {
    lock.releaseLock();
  }
}

// QR印刷ボタン1回につき、対象IDのK列「印刷回数」を1加算する。
// 同じtokenの再送は、通信結果が不明な場合も二重加算しない。
function markPrinted(data) {
  const ids = Array.isArray(data.ids) ? data.ids.map(v => String(v || '').trim()).filter(Boolean) : [];
  const token = String(data.token || '').trim();
  if (!ids.length || ids.length > 20 || new Set(ids).size !== ids.length || ids.some(id => !/^[SOF]\d{8}$/.test(id))) {
    return jsonResponse({ok: false, error: 'INVALID_PRINT_IDS', message: '印刷対象の管理番号を確認してください'});
  }
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(token)) {
    return jsonResponse({ok: false, error: 'INVALID_PRINT_TOKEN', message: '印刷操作を識別できません。QRを作り直してください'});
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const key = 'print_' + token;
    let operation = null;
    const saved = props.getProperty(key);
    if (saved) {
      try { operation = JSON.parse(saved); } catch (error) { throw new Error('PRINT_TOKEN_BROKEN'); }
      if (JSON.stringify(operation.ids) !== JSON.stringify(ids)) throw new Error('PRINT_TOKEN_MISMATCH');
      if (operation.state === 'completed') {
        return jsonResponse({ok: true, apiVersion: 'filament-form-history-v3', action: 'markPrinted', token: token, ids: ids, counts: operation.counts, duplicate: true});
      }
    }

    const book = SpreadsheetApp.getActiveSpreadsheet();
    const ledger = book.getSheetByName('在庫台帳');
    if (!ledger) throw new Error('在庫台帳が見つかりません');
    const headers = ledger.getRange(1, 1, 1, 12).getValues()[0].map(v => String(v || '').trim().normalize('NFKC').replace(/\s/g, ''));
    if (!/印刷回数/.test(headers[10])) throw new Error('在庫台帳K列の見出し「印刷回数」を確認してください');
    const lastRow = ledger.getLastRow();
    if (lastRow < 2) throw new Error('ID_NOT_FOUND');
    const rows = ledger.getRange(2, 1, lastRow - 1, 11).getValues();

    if (!operation) {
      const targets = ids.map(id => {
        const matches = [];
        rows.forEach((row, index) => { if (String(row[0]).trim() === id) matches.push(index); });
        if (matches.length !== 1) throw new Error(matches.length ? 'DUPLICATE_ID:' + id : 'ID_NOT_FOUND:' + id);
        const row = rows[matches[0]], weight = Number(row[6]), raw = row[10];
        if (!Number.isFinite(weight) || weight <= 0) throw new Error('NO_STOCK:' + id);
        const before = raw === '' || raw == null ? 0 : Number(raw);
        if (!Number.isInteger(before) || before < 0) throw new Error('INVALID_PRINT_COUNT:' + id);
        return {id: id, row: matches[0] + 2, before: before, after: before + 1};
      });
      operation = {state: 'pending', ids: ids, targets: targets, createdAt: Date.now()};
      props.setProperty(key, JSON.stringify(operation));
    }

    // pending状態からの再送では、未更新の行だけを書き、更新済みの行はそのままにする。
    operation.targets.forEach(target => {
      const cell = ledger.getRange(target.row, 11), currentRaw = cell.getValue();
      const current = currentRaw === '' || currentRaw == null ? 0 : Number(currentRaw);
      if (current === target.before) cell.setValue(target.after);
      else if (current !== target.after) throw new Error('PRINT_COUNT_REQUIRES_REVIEW:' + target.id);
    });
    SpreadsheetApp.flush();
    const counts = Object.fromEntries(operation.targets.map(target => [target.id, target.after]));
    operation = {...operation, state: 'completed', counts: counts, completedAt: Date.now()};
    props.setProperty(key, JSON.stringify(operation));
    return jsonResponse({ok: true, apiVersion: 'filament-form-history-v3', action: 'markPrinted', token: token, ids: ids, counts: counts, duplicate: false});
  } catch (error) {
    return jsonResponse({ok: false, error: String(error.message), message: '印刷回数を更新できませんでした。印刷せず、もう一度お試しください'});
  } finally {
    lock.releaseLock();
  }
}
// 新窓口からの変更を既存のフォーム回答シートへ同じ12列で追記する。
// 同一GASプロジェクト内の新API更新を直列化する。旧フォーム処理は対象外。
function historyText_(value) { return String(value == null ? '' : value).trim(); }
function historyCell_(value) {
  // 名前・備考などをシート数式として評価させない。
  return typeof value === 'string' && /^[=+@-]/.test(value) ? "'" + value : value;
}
function historySheet_(book) {
  const sheet = book.getSheetByName('フォームの回答 1');
  if (!sheet) throw new Error('「フォームの回答 1」が見つかりません。シート名を確認してください');
  const headers = sheet.getRange(1, 1, 1, 12).getValues()[0].map(v => historyText_(v).normalize('NFKC').replace(/\s/g, ''));
  const tests = [/タイムスタンプ/, /入荷.*使用.*保管/, /管理番号/, /重量/, /^担当$/, /^色$/, /母材/, /サブカテゴリ/, /メーカー/, /保管場所/, /本数/, /備考/];
  if (tests.some((test, i) => !test.test(headers[i]))) throw new Error('フォーム回答の列順が想定と異なります。台帳は更新していません');
  return sheet;
}
function responseNote_(state, operationId, before, after, error) {
  return '[新窓口／' + state + '] 処理ID:' + operationId + '\n' +
    '変更前:' + JSON.stringify(before) + '\n変更後:' + JSON.stringify(after) +
    (error ? '\nエラー:' + error : '');
}
function historyWeight_(value, allowDefault) {
  const empty = value == null || (typeof value === 'string' && !value.trim());
  if (empty && allowDefault) return 1000;
  if (empty || !['number', 'string'].includes(typeof value)) throw new Error('重量がありません');
  const weight = Number(value);
  if (!Number.isFinite(weight) || weight < 0 || weight > 5000) throw new Error('重量は0〜5000gで入力してください');
  return weight;
}
function mutateWithHistory_(data, action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let log = null, logRow = 0, writeAttempted = false, before = null, after = null;
  const operationId = Utilities.getUuid();
  try {
    const id = historyText_(data.id), person = historyText_(data.person);
    if (!id) throw new Error('管理番号がありません');
    if (!person) throw new Error('今回の担当者を選択または入力してください');
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const ledger = book.getSheetByName('在庫台帳');
    if (!ledger) throw new Error('在庫台帳が見つかりません');
    if (action === 'register') {
      const master = requireFilamentMaster_(book, data.filamentKey);
      data.filamentKey = master.filamentKey;
      data.maker = master.maker;
      data.base = master.base;
      data.sub = master.sub;
      data.color = master.color;
      const extraHeaders = ledger.getRange(1, 10, 1, 3).getValues()[0]
        .map(v => historyText_(v).normalize('NFKC').replace(/\s/g, ''));
      if (!/フィラメントキー/.test(extraHeaders[0]) || !/印刷回数/.test(extraHeaders[1]) || !/ラベル注釈/.test(extraHeaders[2])) {
        throw new Error('在庫台帳J～L列の見出しを確認してください');
      }
    }
    const rows = ledger.getLastRow() > 1 ? ledger.getRange(2, 1, ledger.getLastRow() - 1, 9).getValues() : [];
    const indexes = rows.map((r, i) => historyText_(r[0]) === id ? i : -1).filter(i => i >= 0);
    if (indexes.length > 1) throw new Error('DUPLICATE_ID');
    const index = indexes.length ? indexes[0] : -1;
    if (action === 'register' && index >= 0) throw new Error('ID_ALREADY_EXISTS');
    if (action !== 'register' && index < 0) throw new Error('ID_NOT_FOUND');
    before = index >= 0 ? rows[index].slice() : null;

    if (action === 'register') {
      after = [id, historyText_(data.maker), historyText_(data.base), historyText_(data.sub),
        historyText_(data.color), historyText_(data.location), historyWeight_(data.weight, true),
        person, historyText_(data.note)];
    } else {
      after = before.slice();
      after[7] = person;
      if (action === 'updateWeight') after[6] = historyWeight_(data.weight, false);
      else after[5] = historyText_(data.location);
    }
    if (action !== 'updateWeight' && !['さんらいず', 'オーシャン', 'フォージー'].includes(after[5])) {
      throw new Error('保管場所を選択してください');
    }
    log = historySheet_(book);
    const kind = {register: '入荷', updateWeight: '使用', updateLocation: '保管場所移動'}[action];
    const registration = action === 'register';
    const note = responseNote_('処理中', operationId, before, after, '');
    // 旧フォーム同様、使用・移動時の色やメーカーなどは空欄にする。
    const record = [new Date(), kind, id, action === 'updateLocation' ? '' : after[6], person,
      registration ? after[4] : '', registration ? after[2] : '', registration ? after[3] : '',
      registration ? after[1] : '', after[5], registration ? 1 : '',
      (registration && after[8] ? after[8] + '\n' : '') + note];
    // appendRowを使い、通常のフォーム送信の行を上書きしない。
    log.appendRow(record.map(historyCell_));
    SpreadsheetApp.flush();
    const notes = log.getRange(2, 12, log.getLastRow() - 1, 1).getValues();
    const matchingRows = notes.map((r, i) => String(r[0]).includes(operationId) ? i + 2 : -1).filter(i => i > 0);
    if (matchingRows.length !== 1) throw new Error('追記した履歴行を確認できません。台帳は更新していません');
    logRow = matchingRows[0];
    writeAttempted = true;
    if (action === 'register') {
      const ledgerRecord = [...after, historyText_(data.filamentKey), 0, historyText_(data.label)];
      ledger.getRange(ledger.getLastRow() + 1, 1, 1, 12).setValues([ledgerRecord.map(historyCell_)]);
    } else if (action === 'updateWeight') {
      ledger.getRange(index + 2, 7, 1, 2).setValues([[after[6], historyCell_(person)]]);
    } else {
      ledger.getRange(index + 2, 6).setValue(historyCell_(after[5]));
      ledger.getRange(index + 2, 8).setValue(historyCell_(person));
    }
    SpreadsheetApp.flush();
    log.getRange(logRow, 12).setValue(historyCell_((action === 'register' && after[8] ? after[8] + '\n' : '') + responseNote_('完了', operationId, before, after, '')));
    SpreadsheetApp.flush();
    return jsonResponse({ok: true, apiVersion: 'filament-form-history-v3', historyStatus: '完了', historySheet: 'フォームの回答 1', action: action, id: id, weight: after[6], location: after[5], person: person, printCount: action === 'register' ? 0 : undefined, operationId: operationId});
  } catch (error) {
    if (log && logRow) {
      try {
        log.getRange(logRow, 12).setValue(historyCell_(responseNote_(writeAttempted ? '要確認' : '失敗', operationId, before, after, String(error.message))));
        SpreadsheetApp.flush();
      } catch (logError) { console.error(logError); }
    }
    // 台帳とログはトランザクションではない。通信障害時は成功と断定しない。
    return jsonResponse({ok: false, error: writeAttempted ? 'UPDATE_REQUIRES_REVIEW' : String(error.message),
      message: writeAttempted ? 'フォームの回答 1と在庫台帳を確認してください。自動再送しないでください。' : String(error.message), operationId: operationId});
  } finally { lock.releaseLock(); }
}

