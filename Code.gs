/**
 * お店ログ — Google Apps Script バックエンド
 * スプレッドシートに紐づけて使う（拡張機能 → Apps Script）。
 * データは places / visits の2シートに保存される。
 */

const SCHEMA = {
  places: ['id', 'name', 'status', 'genre', 'tags', 'rating', 'memo', 'address', 'lat', 'lng', 'mapsUrl', 'gid', 'createdAt', 'updatedAt'],
  visits: ['id', 'placeId', 'date', 'memo', 'createdAt']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('お店ログ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** GitHub Pages など外部に置いた画面からの呼び出し口。body: {fn, args} */
function doPost(e) {
  let result;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { getAll: getAll, apply: apply, resolveLink: resolveLink };
    if (!fns[req.fn]) throw new Error('unknown function: ' + req.fn);
    result = { ok: true, data: fns[req.fn].apply(null, req.args || []) };
  } catch (err) {
    result = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- データ読み書き ---------- */

function getAll() {
  return {
    places: rows_('places'),
    visits: rows_('visits'),
    appUrl: ScriptApp.getService().getUrl()
  };
}

/** ops: [{op:'upsert', sheet, row} | {op:'delete', sheet, id} | {op:'deletePlace', id}] 再送しても結果が変わらない */
function apply(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ops.forEach(function (op) {
      if (op.op === 'upsert') upsert_(op.sheet, op.row);
      else if (op.op === 'delete') delete_(op.sheet, 'id', op.id);
      else if (op.op === 'deletePlace') {
        delete_('places', 'id', op.id);
        delete_('visits', 'placeId', op.id);
      }
    });
  } finally {
    lock.releaseLock();
  }
  return true;
}

function sheet_(name) {
  if (!SCHEMA[name]) throw new Error('unknown sheet: ' + name);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    // 日付や数値に自動変換されないよう全列を書式なしテキストにする
    sh.getRange(1, 1, sh.getMaxRows(), SCHEMA[name].length).setNumberFormat('@');
    sh.appendRow(SCHEMA[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(name) {
  const cols = SCHEMA[name];
  const values = sheet_(name).getDataRange().getDisplayValues();
  return values.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    cols.forEach(function (c, i) { o[c] = r[i] || ''; });
    return o;
  });
}

function cell_(v) {
  const s = v == null ? '' : String(v);
  return /^[=+]/.test(s) ? "'" + s : s; // 数式として解釈させない
}

function upsert_(name, row) {
  const sh = sheet_(name);
  const cols = SCHEMA[name];
  const vals = cols.map(function (c) { return cell_(row[c]); });
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues().map(function (r) { return r[0]; });
  const i = ids.indexOf(String(row.id));
  if (i > 0) sh.getRange(i + 1, 1, 1, cols.length).setValues([vals]);
  else sh.appendRow(vals);
}

function delete_(name, col, value) {
  const sh = sheet_(name);
  const c = SCHEMA[name].indexOf(col) + 1;
  const v = sh.getRange(1, c, sh.getLastRow(), 1).getDisplayValues();
  for (let r = v.length - 1; r >= 1; r--) {
    if (v[r][0] === String(value)) sh.deleteRow(r + 1);
  }
}

/* ---------- Googleマップのリンク解析 ---------- */

/** 共有リンク（短縮URL可）や共有テキストから店名・住所・座標を取り出す */
function resolveLink(input) {
  const text = String(input || '').trim();
  const m = text.match(/https?:\/\/[^\s]+/);
  if (!m) return { name: text.split('\n')[0] };

  const original = m[0];
  let url = original;
  // maps.app.goo.gl などの短縮URLはリダイレクト先を辿る
  for (let i = 0; i < 5 && !isFullMapsUrl_(url); i++) {
    const res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    const headers = res.getHeaders();
    let loc = headers['Location'] || headers['location'];
    if (!loc) break;
    const cont = loc.match(/consent\.google\.[^/]+\/.*[?&]continue=([^&]+)/);
    if (cont) loc = decodeURIComponent(cont[1]);
    url = loc;
  }

  const out = parseMapsUrl(url);
  out.mapsUrl = original;

  // URLに座標や住所が無いときは Maps サービス（無料・キー不要）で補う
  try {
    const geocoder = Maps.newGeocoder().setLanguage('ja').setRegion('jp');
    if (out.lat == null && (out.query || out.name)) {
      const r = geocoder.geocode(out.query || out.name).results[0];
      if (r) {
        out.lat = r.geometry.location.lat;
        out.lng = r.geometry.location.lng;
        out.address = out.address || cleanAddress_(r.formatted_address);
      }
    } else if (out.lat != null && !out.address) {
      const r = geocoder.reverseGeocode(out.lat, out.lng).results[0];
      if (r) out.address = cleanAddress_(r.formatted_address);
    }
  } catch (e) {
    // 住所・座標は任意項目なので失敗しても登録は続行できる
  }
  delete out.query;
  return out;
}

function isFullMapsUrl_(url) {
  return /\/maps\/(place|search)\/|[?&](q|query)=/.test(url);
}

function cleanAddress_(a) {
  return String(a || '').replace(/^日本、\s*/, '');
}

/** index.html にも同じ関数がある（ローカル動作用）。変更時は両方直すこと */
function parseMapsUrl(url) {
  const out = {};
  const dec = function (s) {
    try { return decodeURIComponent(s.replace(/\+/g, ' ')); } catch (e) { return s; }
  };
  let m = url.match(/\/maps\/(?:place|search)\/([^/@?]+)/);
  if (m) out.name = dec(m[1]);
  m = url.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) { out.lat = Number(m[1]); out.lng = Number(m[2]); }
  m = url.match(/0x[0-9a-f]+:0x[0-9a-f]+/i);
  if (m) out.gid = m[0].toLowerCase();
  m = url.match(/[?&](?:q|query)=([^&#]+)/);
  if (m) {
    const q = dec(m[1]);
    const ll = q.match(/^\s*(-?\d+\.\d+),\s*(-?\d+\.\d+)\s*$/);
    if (ll) {
      if (out.lat == null) { out.lat = Number(ll[1]); out.lng = Number(ll[2]); }
    } else {
      out.query = q;
      // iPhoneの共有は「店名 住所」が1つになっていることがある
      const parts = q.split(/\n|,\s*|、|\s*〒/);
      if (!out.name) out.name = parts[0].trim();
      if (q.length > parts[0].length) out.address = q.slice(parts[0].length).replace(/^[\s,、]+/, '').trim();
    }
  }
  return out;
}
