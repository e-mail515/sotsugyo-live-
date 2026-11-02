/**
 * 卒業ライブ バンドエントリー一覧 API（回答スプレッドシートにバインドして使う）
 *
 * 【初回デプロイ手順】
 * 1. 回答スプレッドシートを開き「拡張機能 → Apps Script」を選ぶ
 * 2. このファイルの中身を Code.gs に丸ごと貼り付けて保存
 * 3. 上部の関数選択で「testOutput」を選んで実行（初回は権限の承認画面が出るので許可）
 *    → 実行ログにJSONが出ること・メールアドレスが含まれないことを確認
 * 4. 右上「デプロイ → 新しいデプロイ」→ 種類の選択（歯車）で「ウェブアプリ」を選ぶ
 *      次のユーザーとして実行：自分
 *      アクセスできるユーザー：全員
 * 5. 「デプロイ」を押し、表示された「ウェブアプリのURL」（…/exec で終わる）を
 *    docs/index.html の API_URL に貼る
 *
 * 【コードを書き換えたとき】
 * 保存しただけでは公開中のURLに反映されない。
 * 「デプロイ → デプロイを管理 → 鉛筆（編集）→ バージョン：新バージョン → デプロイ」を行う。
 * ※「新しいデプロイ」を作り直すとURLが変わるので、サイト側の API_URL も直す必要がある。
 *
 * 【レスポンス形式】（軽量化のため短いキー・インデントなし）
 *   {"v":"ハッシュ","n":件数,"d":["3/11(木)",...],"b":[バンド,...]}
 *   ?v= に現在のハッシュを付けて呼ばれ、データに変化がなければ {"nm":1} だけ返す。
 *   バンドのキーは下の COLUMNS のとおり。値が空のキーは省略する。
 *   u = 不可時間。d と同じ順の配列の配列（例 [["9:00~","9:30~"],[],[],["13:00~"]]）
 */

// ヘッダー名 → 出力キー。ここに書いた列だけを出力する（ホワイトリスト）。
// フォームの質問文を変えたら、ここのヘッダー名も合わせて書き換える。
const COLUMNS = {
  'バンド形態': 'f',
  'コピーするバンド': 'c',
  'バンドリーダー(名前)': 'l',
  'ボーカル(名前)': 'vo',
  'ボーカルのパート': 'vp',
  'ギター1(名前)': 'g1',
  'ギター2(名前)': 'g2',
  'ベース(名前)': 'ba',
  'ドラム(名前)': 'dr',
  'キーボード(名前)': 'kb',
  'その他': 'ot',
};
// 「不可時間 [3/11(木)]」形式の列。[ ] の中が日付ラベルになる（正規化後の文字列に対して判定）
const UNAVAILABLE_RE = /^不可時間\[(.+)\]$/;
// 並び順に使うだけで出力はしない
const TIMESTAMP_HEADER = 'タイムスタンプ';
// 回答シート名。空なら先頭のシート
const SHEET_NAME = '';
const CACHE_SEC = 30;

function doGet(e) {
  const data = getData_();
  const v = e && e.parameter && e.parameter.v;
  return ContentService.createTextOutput(v && v === data.v ? '{"nm":1}' : data.json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getData_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.getAll(['json', 'v']);
  if (hit.json && hit.v) return { json: hit.json, v: hit.v };
  const data = build_(readSheet_());
  try {
    cache.putAll(data, CACHE_SEC);
  } catch (err) {
    // キャッシュの上限（1件100KB）を超えても、レスポンスはそのまま返す
  }
  return data;
}

function readSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  return sheet.getDataRange().getDisplayValues();
}

// 全角/半角・空白の違いでヘッダーが一致しなくならないように正規化する
function norm_(s) {
  return String(s).normalize('NFKC').replace(/\s+/g, '');
}

function build_(rows) {
  const whitelist = {};
  Object.keys(COLUMNS).forEach(h => { whitelist[norm_(h)] = COLUMNS[h]; });

  const cols = []; // [列番号, キー]
  const days = []; // [列番号, 日付ラベル]
  let ts = -1;
  rows[0].map(norm_).forEach((h, i) => {
    if (whitelist[h]) cols.push([i, whitelist[h]]);
    else if (h === norm_(TIMESTAMP_HEADER)) ts = i;
    else {
      const m = h.match(UNAVAILABLE_RE);
      if (m) days.push([i, m[1]]);
    }
  });

  const list = [];
  rows.slice(1).forEach((r, n) => {
    const cell = i => String(r[i] == null ? '' : r[i]).trim();
    const b = {};
    cols.forEach(([i, k]) => {
      const s = cell(i);
      if (s) b[k] = s;
    });
    const u = days.map(([i]) => cell(i).split(',').map(s => s.trim()).filter(Boolean));
    if (!Object.keys(b).length && !u.some(a => a.length)) return; // 完全に空の行
    b.u = u;
    const t = ts < 0 ? NaN : Date.parse(cell(ts));
    list.push({ b: b, t: isNaN(t) ? Infinity : t, n: n });
  });
  // タイムスタンプ順。読めない場合はシートの行順
  list.sort((a, b) => (a.t - b.t) || (a.n - b.n));

  const d = days.map(x => x[1]);
  const bands = list.map(x => x.b);
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, JSON.stringify({ d: d, b: bands }), Utilities.Charset.UTF_8);
  const v = Utilities.base64EncodeWebSafe(digest).slice(0, 11);
  return { v: v, json: JSON.stringify({ v: v, n: bands.length, d: d, b: bands }) };
}

// サンプル行と実データでJSONを組み立ててログに出す。メールアドレスが混ざっていたらエラーにする。
function testOutput() {
  const header = ['タイムスタンプ', 'メールアドレス', 'バンド形態', 'コピーするバンド', 'バンドリーダー(名前)',
    'ボーカル(名前)', 'ボーカルのパート', 'ギター1(名前)', 'ギター2(名前)', 'ベース(名前)', 'ドラム(名前)',
    'キーボード(名前)', 'その他', '不可時間 [3/11(木)]', '不可時間 [3/12(金)]', '不可時間 [3/13(土)]',
    '不可時間 [3/14(日)]', '演奏曲', '対バン(任意）', 'その他なにかあれば自由に記入してください。'];
  const sample = ['2026/11/02 23:28:29', 'test@example.com', '通常バンド', 'あああああああああああ', 'あ',
    'あ', 'ピン', 'い', 'う', 'え', 'お', 'か', 'き(トランペット)', '9:00~, 9:30~, 11:00~', '9:00~, 11:00~',
    '9:30~, 11:00~, 12:00~', '11:30~, 13:00~', 'きくけこ', '', ''];
  const earlier = ['2026/11/01 10:00:00', 'early@example.com', '通常バンド', '先に送信したバンド', 'さ',
    'さ', 'Gt.Vo', '', '', 'し', 'す', '', '', '', '', '', '', '曲', '', 'メモ'];
  const blank = header.map(() => '');

  const data = build_([header, sample, blank, earlier]);
  Logger.log(data.json);
  if (/@|きくけこ|メモ|2026\/11/.test(data.json)) throw new Error('出力してはいけない列が含まれています');

  let sheetRows;
  try {
    sheetRows = readSheet_();
  } catch (err) {
    Logger.log('スプレッドシートにバインドされていないため、実データの確認はスキップしました');
    return;
  }
  const real = build_(sheetRows);
  Logger.log('実データ（' + JSON.parse(real.json).n + '件）: ' + real.json);
  if (real.json.indexOf('@') >= 0) Logger.log('⚠ 出力に「@」が含まれています。メールアドレスが混ざっていないか確認してください');
}
