/**
 * 戰勝21點 · 練習記錄後端
 * 綁定在「戰勝21點 練習記錄」試算表上，由 GitHub Pages 上的前端呼叫。
 *
 * 帳號     ：暱稱 + 四位數字密碼（只存 SHA-256 雜湊，試算表看不到原始密碼）
 * 手牌記錄 ：每一次判斷一列（策略判斷與算牌題目都寫這裡，用「來源」欄區分）
 * 每日統計 ：日期 × 練習者，策略手數與算牌題數分開兩組欄位（增量更新，不重算全表）
 * 登入記錄 ：每次登入一列，給系統人員的使用者儀表板看
 *
 * 寫入走 POST（跨網域 no-cors，前端不需要讀回應）
 * 讀取走 GET + JSONP（跨網域要讀回應，只能用 JSONP）
 */

var SH_HANDS = '手牌記錄';
var SH_DAILY = '每日統計';
var SH_USERS = '帳號';
var SH_LOGIN = '登入記錄';
var HEAD_HANDS = ['時間', '日期', '練習者', '來源', '手牌', '莊家明牌', '你選', '正解', '對錯', '秒數'];
var HEAD_DAILY = ['日期', '練習者', '手數', '正確', '正確率', '算牌題數', '算牌正確'];
var HEAD_USERS = ['暱稱', '密碼雜湊', '建立時間', '最後使用', '角色', '登入次數', '最後登入', 'Google ID'];
var HEAD_LOGIN = ['時間', '暱稱', '角色'];

var SRC_COUNT = '算牌';     // 手牌記錄的「來源」欄是這個值，就計入算牌題數而不是策略手數
var LOGIN_KEEP = 2000;      // 登入記錄只留最近這麼多列
// 使用者都在台灣，日界線就該用台北時間切。
// 試算表本身的時區是 America/Los_Angeles，跟著它走的話下午三點前練的會被記到前一天。
var TZ = 'Asia/Taipei';

// 帳號表欄位位置（1-based），改欄位順序時這裡要一起改
var U_NAME = 1, U_HASH = 2, U_CREATED = 3, U_USED = 4, U_ROLE = 5, U_LOGINS = 6, U_LASTLOGIN = 7,
    U_GID = 8;

// Google 登入用的 OAuth 用戶端 ID。這不是密鑰，它本來就會出現在前端網頁原始碼裡。
// 空字串＝Google 登入關閉，前端不會顯示那顆按鈕。
var GOOGLE_CLIENT_ID = '';

/**
 * 放在最前面：編輯器預設會選第一個函式，所以第一個函式必須無參數且可安全執行。
 * 建立工作表、補上後來新增的欄位、移除預設空白表。重複執行不會有副作用。
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetOf(ss, SH_HANDS, HEAD_HANDS);
  sheetOf(ss, SH_DAILY, HEAD_DAILY);
  sheetOf(ss, SH_USERS, HEAD_USERS);
  sheetOf(ss, SH_LOGIN, HEAD_LOGIN);
  var keep = [SH_HANDS, SH_DAILY, SH_USERS, SH_LOGIN];
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (keep.indexOf(all[i].getName()) < 0 && all[i].getLastRow() === 0) {
      ss.deleteSheet(all[i]);
    }
  }
  return '完成：' + ss.getSheets().map(function (s) { return s.getName(); }).join(' / ');
}

function ymd(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

/**
 * 時間欄一律寫成台北時間的純文字。
 *
 * 原因：把 Date 寫進儲存格，Sheets 會存成「試算表時區的牆上時間」，
 * 讀回來卻是把那串牆上時間當成 UTC 的 Date —— 整個被平移了試算表的時區偏移
 * （這張表是 America/Los_Angeles，實測差 7 小時）。存文字就完全繞開這件事。
 */
function nowText() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

/** 試算表時區相對 UTC 的偏移量，用來還原舊資料裡被平移過的 Date */
function sheetOffsetMs(ss) {
  var t = new Date();
  var utc = Utilities.formatDate(t, 'UTC', 'yyyy/MM/dd HH:mm:ss');
  var loc = Utilities.formatDate(t, ss.getSpreadsheetTimeZone() || TZ, 'yyyy/MM/dd HH:mm:ss');
  return new Date(loc).getTime() - new Date(utc).getTime();
}

/**
 * 還原成真正的時間點。舊資料是 Date（被平移過，補回偏移量）；新資料是文字（直接解析）。
 * 註：補償用的是「現在」的偏移量，跨日光節約時間的舊資料可能差一小時，
 * 但只影響顯示與極少數剛好卡在午夜的分日，新資料存文字後不會再有這問題。
 */
function realDate(v, off) {
  if (v instanceof Date) return new Date(v.getTime() - off);
  var s = String(v || '').trim();
  if (!s) return null;
  var d = new Date(s.indexOf('T') < 0 ? s.replace(/-/g, '/') : s);
  return isNaN(d.getTime()) ? null : d;
}

function stampOf(v, off) {
  var d = realDate(v, off);
  return d ? Utilities.formatDate(d, TZ, 'MM/dd HH:mm') : (v ? String(v) : '');
}

function ymdOf(v, off) {
  var d = realDate(v, off);
  return d ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : '';
}

function sheetOf(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    writeHead(sh, head);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, head.length);
    // 時間/日期欄一開始就設成純文字，否則 Sheets 會轉成日期型別，再被試算表時區位移一天。
    //
    // 只在「建表當下」設，絕對不要對既有資料補設：
    // 把已經存著日期值的欄位改成純文字，getValues() 會只讀到顯示字串
    // （顯示格式是只有日期的話，時分秒就直接消失）。這件事踩過一次，不要再踩。
    if (name === SH_DAILY) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
    if (name === SH_HANDS) sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@');
    if (name === SH_USERS) sh.getRange(1, U_CREATED, sh.getMaxRows(), 2).setNumberFormat('@');
    if (name === SH_LOGIN) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  } else if (sh.getLastColumn() < head.length) {
    // 舊表補欄位：只寫標題，既有資料列留空，讀取一律用 || 0 兜底
    writeHead(sh, head);
  }
  return sh;
}

function writeHead(sh, head) {
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1E2327').setFontColor('#EDEBE6');
}

function out(obj, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════ 帳號 ═══════════ */

function hashPin(name, pin) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, 'bj21|' + name + '|' + pin, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function cleanName(v) {
  return String(v || '').trim().substring(0, 20);
}

function findUser(ss, name) {
  var sh = sheetOf(ss, SH_USERS, HEAD_USERS);
  var last = sh.getLastRow();
  if (last < 2) return { sh: sh, row: -1 };
  var v = sh.getRange(2, 1, last - 1, HEAD_USERS.length).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === name) {
      return {
        sh: sh, row: i + 2,
        hash: String(v[i][U_HASH - 1]),
        role: String(v[i][U_ROLE - 1] || 'user'),
        logins: Number(v[i][U_LOGINS - 1]) || 0,
        gid: String(v[i][U_GID - 1] || '')
      };
    }
  }
  return { sh: sh, row: -1 };
}

/** 四位數字密碼，或 Google 登入發的 64 碼權杖。兩者都只存雜湊，驗證方式一模一樣。 */
function isSecret(s) {
  return /^\d{4}$/.test(s) || /^[0-9a-f]{64}$/.test(s);
}

function newToken() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

/**
 * 驗證帳號。createIfMissing=true 時，暱稱不存在就直接建立（等於註冊）。
 * 回傳 {ok, created, name, role, row, sh, logins, gid}
 * 或 {ok:false, error:'bad_input'|'no_user'|'bad_pin'}
 */
function auth(ss, name, pin, createIfMissing) {
  name = cleanName(name);
  pin = String(pin || '');
  if (!name || !isSecret(pin)) return { ok: false, error: 'bad_input' };
  var u = findUser(ss, name);
  var h = hashPin(name, pin);
  if (u.row < 0) {
    if (!createIfMissing) return { ok: false, error: 'no_user' };
    u.sh.appendRow([name, h, nowText(), nowText(), 'user', 0, '', '']);
    return { ok: true, created: true, name: name, role: 'user', row: u.sh.getLastRow(), sh: u.sh, logins: 0 };
  }
  if (u.hash !== h) return { ok: false, error: 'bad_pin' };
  u.sh.getRange(u.row, U_USED).setValue(nowText());
  return { ok: true, created: false, name: name, role: u.role, row: u.row, sh: u.sh,
           logins: u.logins, gid: u.gid };
}

/* ═══════════ Google 登入 ═══════════ */

/**
 * 驗證前端送來的 Google ID Token。
 *
 * 一定要在後端驗：前端自己解出來的 sub 誰都能偽造。
 * 用 tokeninfo 端點驗，不必自己處理 JWT 簽章與金鑰輪替。
 * ⚠️ 這支用到 UrlFetchApp，所以指令碼需要 script.external_request 權限，
 *    加上這個權限之後擁有者必須重新授權一次，否則 /exec 會回授權錯誤。
 */
function verifyGoogle(idt) {
  if (!GOOGLE_CLIENT_ID) return { ok: false, error: 'google_off' };
  idt = String(idt || '');
  if (idt.length < 20 || idt.length > 4000) return { ok: false, error: 'bad_token' };
  var res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idt),
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { ok: false, error: 'bad_token' };
  var d;
  try { d = JSON.parse(res.getContentText()); } catch (e) { return { ok: false, error: 'bad_token' }; }
  if (String(d.aud) !== GOOGLE_CLIENT_ID) return { ok: false, error: 'bad_aud' };
  if (Number(d.exp) * 1000 < new Date().getTime()) return { ok: false, error: 'expired' };
  if (!d.sub) return { ok: false, error: 'bad_token' };
  var sug = String(d.given_name || d.name || String(d.email || '').split('@')[0] || '').trim();
  return { ok: true, sub: String(d.sub), suggest: sug.substring(0, 12) };
}

function findByGoogle(ss, sub) {
  var sh = sheetOf(ss, SH_USERS, HEAD_USERS);
  var last = sh.getLastRow();
  if (last < 2) return { sh: sh, row: -1 };
  var v = sh.getRange(2, 1, last - 1, HEAD_USERS.length).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][U_GID - 1]) === sub && sub) {
      return {
        sh: sh, row: i + 2, name: String(v[i][0]),
        role: String(v[i][U_ROLE - 1] || 'user'),
        logins: Number(v[i][U_LOGINS - 1]) || 0
      };
    }
  }
  return { sh: sh, row: -1 };
}

/** 發一把新權杖給這一列並回傳明文（只有這一次看得到，之後表上只有雜湊） */
function issueToken(sh, row, name) {
  var tok = newToken();
  sh.getRange(row, U_HASH).setValue(hashPin(name, tok));
  return tok;
}

/** 管理動作共用：驗身分 + 檢查角色。前端的角色只是快取，這裡才是真的關卡。 */
function requireAdmin(ss, p) {
  var a = auth(ss, p.name, p.pin, false);
  if (!a.ok) return a;
  if (a.role !== 'admin') return { ok: false, error: 'forbidden' };
  return a;
}

function hasAnyAdmin(ss) {
  var sh = sheetOf(ss, SH_USERS, HEAD_USERS);
  var last = sh.getLastRow();
  if (last < 2) return false;
  var v = sh.getRange(2, U_ROLE, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === 'admin') return true;
  }
  return false;
}

/** 記一次登入：帳號表累加次數，另外在登入記錄留一列時間 */
function markLogin(ss, a) {
  a.sh.getRange(a.row, U_LOGINS).setValue((a.logins || 0) + 1);
  a.sh.getRange(a.row, U_LASTLOGIN).setValue(nowText());
  var sh = sheetOf(ss, SH_LOGIN, HEAD_LOGIN);
  sh.appendRow([nowText(), a.name, a.role || 'user']);
  var last = sh.getLastRow();
  if (last > LOGIN_KEEP + 500) sh.deleteRows(2, last - 1 - LOGIN_KEEP);
}

/** 取某個練習者的每日統計；策略手數與算牌題數分開回傳 */
function daysOf(ss, name) {
  var sh = ss.getSheetByName(SH_DAILY);
  if (!sh || sh.getLastRow() < 2) return { days: [], count: [] };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_DAILY.length).getValues();
  var days = [], count = [];
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][1]) !== name) continue;
    var day = ymd(v[i][0]);
    var hands = Number(v[i][2]) || 0;
    var cq = Number(v[i][5]) || 0;
    if (hands) days.push({ day: day, hands: hands, correct: Number(v[i][3]) || 0 });
    if (cq) count.push({ day: day, q: cq, c: Number(v[i][6]) || 0 });
  }
  return { days: days, count: count };
}

/** 系統人員儀表板：每個帳號的登入與練習狀況 */
function adminUsers(ss) {
  var sh = sheetOf(ss, SH_USERS, HEAD_USERS);
  var off = sheetOffsetMs(ss);
  var last = sh.getLastRow();
  if (last < 2) return [];

  // 每日統計整張讀一次在記憶體裡彙總，不要每個帳號各掃一遍
  var agg = {}, today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var d = ss.getSheetByName(SH_DAILY);
  if (d && d.getLastRow() > 1) {
    var dv = d.getRange(2, 1, d.getLastRow() - 1, HEAD_DAILY.length).getValues();
    for (var i = 0; i < dv.length; i++) {
      var nm = String(dv[i][1]);
      if (!nm) continue;
      var day = ymd(dv[i][0]);
      var g = agg[nm];
      if (!g) g = agg[nm] = { hands: 0, correct: 0, count: 0, ccorrect: 0, days: 0, lastDay: '', todayHands: 0, todayCount: 0 };
      g.hands += Number(dv[i][2]) || 0;
      g.correct += Number(dv[i][3]) || 0;
      g.count += Number(dv[i][5]) || 0;
      g.ccorrect += Number(dv[i][6]) || 0;
      g.days++;
      if (day > g.lastDay) g.lastDay = day;
      if (day === today) {
        g.todayHands += Number(dv[i][2]) || 0;
        g.todayCount += Number(dv[i][5]) || 0;
      }
    }
  }

  var v = sh.getRange(2, 1, last - 1, HEAD_USERS.length).getValues(), res = [];
  for (var j = 0; j < v.length; j++) {
    var name = String(v[j][0]);
    if (!name) continue;
    var g2 = agg[name] || { hands: 0, correct: 0, count: 0, ccorrect: 0, days: 0, lastDay: '', todayHands: 0, todayCount: 0 };
    res.push({
      name: name,
      role: String(v[j][U_ROLE - 1] || 'user'),
      created: stampOf(v[j][U_CREATED - 1], off),
      // 舊帳號沒有「最後登入」，退回用「最後使用」
      lastLogin: stampOf(v[j][U_LASTLOGIN - 1] || v[j][U_USED - 1], off),
      logins: Number(v[j][U_LOGINS - 1]) || 0,
      google: !!String(v[j][U_GID - 1] || ''),
      hands: g2.hands, correct: g2.correct,
      count: g2.count, ccorrect: g2.ccorrect,
      days: g2.days, lastDay: g2.lastDay,
      todayHands: g2.todayHands, todayCount: g2.todayCount
    });
  }
  res.sort(function (a, b) { return String(b.lastLogin).localeCompare(String(a.lastLogin)); });
  return res;
}

/* ═══════════ 寫入 ═══════════ */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return out({ ok: false, error: 'busy' });
  }
  try {
    var body = JSON.parse(e.postData.contents);
    var rows = body.rows || [];
    if (!rows.length) return out({ ok: true, n: 0 });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var a = auth(ss, body.player, body.pin, false);
    if (!a.ok) return out({ ok: false, error: a.error });
    var player = a.name;

    var hands = sheetOf(ss, SH_HANDS, HEAD_HANDS);
  
    // rows: [isoTime, 來源, 手牌, 莊家明牌, 你選, 正解, O|X, 秒數]
    var buf = [], daily = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var when = new Date(r[0]);
      if (isNaN(when.getTime())) when = new Date();
      var day = Utilities.formatDate(when, TZ, 'yyyy-MM-dd');
      buf.push([Utilities.formatDate(when, TZ, 'yyyy-MM-dd HH:mm:ss'), day, player,
                r[1], r[2], r[3], r[4], r[5], r[6], r[7]]);
      var k = day + ' ' + player;
      if (!daily[k]) daily[k] = { day: day, player: player, n: 0, ok: 0, cn: 0, cok: 0 };
      // 算牌題目不能混進策略手數，否則前端的「練習記錄」正確率會被稀釋
      if (r[1] === SRC_COUNT) {
        daily[k].cn++;
        if (r[6] === 'O') daily[k].cok++;
      } else {
        daily[k].n++;
        if (r[6] === 'O') daily[k].ok++;
      }
    }
    hands.getRange(hands.getLastRow() + 1, 1, buf.length, HEAD_HANDS.length).setValues(buf);
    bumpDaily(ss, daily);
    return out({ ok: true, n: buf.length });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 增量更新每日統計：只讀現有列找 key，不重算整張手牌記錄 */
function bumpDaily(ss, daily) {
  var sh = sheetOf(ss, SH_DAILY, HEAD_DAILY);
  var last = sh.getLastRow();
  var idx = {}, vals = [];
  if (last > 1) {
    vals = sh.getRange(2, 1, last - 1, HEAD_DAILY.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      idx[ymd(vals[i][0]) + ' ' + String(vals[i][1])] = i;
    }
  }
  var appends = [];
  for (var k in daily) {
    var e = daily[k];
    if (idx.hasOwnProperty(k)) {
      var i2 = idx[k];
      var n = (Number(vals[i2][2]) || 0) + e.n;
      var ok = (Number(vals[i2][3]) || 0) + e.ok;
      var cn = (Number(vals[i2][5]) || 0) + e.cn;
      var cok = (Number(vals[i2][6]) || 0) + e.cok;
      sh.getRange(i2 + 2, 3, 1, 5).setValues([[n, ok, n ? ok / n : 0, cn, cok]]);
    } else {
      appends.push([e.day, e.player, e.n, e.ok, e.n ? e.ok / e.n : 0, e.cn, e.cok]);
    }
  }
  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, HEAD_DAILY.length).setValues(appends);
  }
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 5, sh.getLastRow() - 1, 1).setNumberFormat('0.0%');
  }
}

/**
 * 用手牌記錄的絕對時間，照台北時區把每日統計整張重算。
 * 之前每日統計是照試算表時區（美西）切天的，日期整批往前偏一天；
 * 手牌記錄第一欄存的是真正的時間點，所以可以完全重建，不是估的。
 */
function rebuildDaily(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var hands = sheetOf(ss, SH_HANDS, HEAD_HANDS);
  var off = sheetOffsetMs(ss);
  var last = hands.getLastRow();
  var daily = {}, order = [];
  if (last > 1) {
    var v = hands.getRange(2, 1, last - 1, HEAD_HANDS.length).getValues();
    var fixed = [];
    for (var i = 0; i < v.length; i++) {
      var player = String(v[i][2]);
      var old = String(v[i][1] || '');
      var real = realDate(v[i][0], off);
      // 順便把時間欄從 Date 改寫成台北時間的純文字，之後就不必再靠偏移量還原
      var timeText = real ? Utilities.formatDate(real, TZ, 'yyyy-MM-dd HH:mm:ss') : String(v[i][0] || '');
      if (!player) { fixed.push([timeText, old]); continue; }
      // 時間欄解不出來才退回用原本的日期字串，寧可日期偏一天也不要整列漏掉
      var day = real ? Utilities.formatDate(real, TZ, 'yyyy-MM-dd') : old;
      fixed.push([timeText, day]);
      if (!day) continue;
      var k = day + ' ' + player;
      var e = daily[k];
      if (!e) { e = daily[k] = { day: day, player: player, n: 0, ok: 0, cn: 0, cok: 0 }; order.push(k); }
      if (String(v[i][3]) === SRC_COUNT) {
        e.cn++; if (v[i][8] === 'O') e.cok++;
      } else {
        e.n++; if (v[i][8] === 'O') e.ok++;
      }
    }
    hands.getRange(2, 1, fixed.length, 2).setValues(fixed);
  }

  var sh = sheetOf(ss, SH_DAILY, HEAD_DAILY);
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  order.sort();
  var rows = order.map(function (k) {
    var e = daily[k];
    return [e.day, e.player, e.n, e.ok, e.n ? e.ok / e.n : 0, e.cn, e.cok];
  });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, HEAD_DAILY.length).setValues(rows);
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('0.0%');
  }
  return rows.length;
}

/**
 * 清掉某個人的算牌記錄：手牌記錄刪掉他來源是「算牌」的列，每日統計的算牌兩欄歸零。
 * 策略練習的手數完全不動。手牌記錄也真的刪，否則 rebuildDaily 會把數字算回來。
 */
function clearCount(ss, name) {
  var hands = sheetOf(ss, SH_HANDS, HEAD_HANDS);
  var last = hands.getLastRow(), removed = 0;
  if (last > 1) {
    var v = hands.getRange(2, 1, last - 1, HEAD_HANDS.length).getValues();
    var keep = [];
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][2]) === name && String(v[i][3]) === SRC_COUNT) { removed++; continue; }
      keep.push(v[i]);
    }
    if (removed) {
      hands.getRange(2, 1, last - 1, HEAD_HANDS.length).clearContent();
      if (keep.length) hands.getRange(2, 1, keep.length, HEAD_HANDS.length).setValues(keep);
    }
  }

  var d = sheetOf(ss, SH_DAILY, HEAD_DAILY);
  var dl = d.getLastRow();
  if (dl > 1) {
    var dv = d.getRange(2, 1, dl - 1, HEAD_DAILY.length).getValues(), touched = false;
    for (var j = 0; j < dv.length; j++) {
      if (String(dv[j][1]) !== name) continue;
      if ((Number(dv[j][5]) || 0) || (Number(dv[j][6]) || 0)) { dv[j][5] = 0; dv[j][6] = 0; touched = true; }
    }
    if (touched) d.getRange(2, 1, dl - 1, HEAD_DAILY.length).setValues(dv);
  }
  return removed;
}

/* ═══════════ 讀取（JSONP）═══════════ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (p.action === 'login' || p.action === 'load' || p.action === 'register') {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(20000); } catch (err) { return out({ ok: false, error: 'busy' }, cb); }
    try {
      var isReg = p.action === 'register';
      // 註冊才建帳號。登入不建：打錯字應該要報「沒有這個帳號」，
      // 不是默默開一個新的、讓人以為練習記錄不見了。
      if (isReg && findUser(ss, cleanName(p.name)).row > 0) {
        return out({ ok: false, error: 'exists' }, cb);
      }
      var a = auth(ss, p.name, p.pin, isReg);
      if (!a.ok) return out({ ok: false, error: a.error }, cb);
      if (p.action !== 'load') markLogin(ss, a);
      var st = daysOf(ss, a.name);
      return out({
        ok: true, created: !!a.created, name: a.name, role: a.role,
        days: st.days, count: st.count
      }, cb);
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * 第一位系統人員的開通入口：已經有系統人員之後就永遠回 already。
   * 因為自我關閉，把它留在公開原始碼裡不會變成提權管道。
   */
  if (p.action === 'bootstrap_admin') {
    var lk = LockService.getScriptLock();
    try { lk.waitLock(20000); } catch (err) { return out({ ok: false, error: 'busy' }, cb); }
    try {
      if (hasAnyAdmin(ss)) return out({ ok: false, error: 'already' }, cb);
      var b = auth(ss, p.name, p.pin, true);
      if (!b.ok) return out({ ok: false, error: b.error }, cb);
      b.sh.getRange(b.row, U_ROLE).setValue('admin');
      return out({ ok: true, name: b.name, role: 'admin', created: !!b.created }, cb);
    } finally {
      lk.releaseLock();
    }
  }

  if (p.action === 'admin_users') {
    var au = requireAdmin(ss, p);
    if (!au.ok) return out(au, cb);
    return out({ ok: true, users: adminUsers(ss) }, cb);
  }

  if (p.action === 'admin_create') {
    var ac = requireAdmin(ss, p);
    if (!ac.ok) return out(ac, cb);
    var nn = cleanName(p.newname), np = String(p.newpin || '');
    if (!nn || !/^\d{4}$/.test(np)) return out({ ok: false, error: 'bad_input' }, cb);
    var ex = findUser(ss, nn);
    if (ex.row > 0) return out({ ok: false, error: 'exists' }, cb);
    ex.sh.appendRow([nn, hashPin(nn, np), nowText(), nowText(), 'user', 0, '', '']);
    return out({ ok: true, name: nn }, cb);
  }

  if (p.action === 'admin_setpin') {
    var ap = requireAdmin(ss, p);
    if (!ap.ok) return out(ap, cb);
    var tn = cleanName(p.target), tp = String(p.newpin || '');
    if (!tn || !/^\d{4}$/.test(tp)) return out({ ok: false, error: 'bad_input' }, cb);
    var t = findUser(ss, tn);
    if (t.row < 0) return out({ ok: false, error: 'no_user' }, cb);
    t.sh.getRange(t.row, U_HASH).setValue(hashPin(tn, tp));
    return out({ ok: true, name: tn }, cb);
  }

  /**
   * Google 登入。三種情況一支處理完：
   *   已綁過      → 直接登入，發一把新權杖
   *   沒綁過＋name→ 開新帳號
   *   沒綁過＋link→ 把這個 Google 帳號綁到既有的暱稱＋密碼帳號（記錄就跟著過來）
   *   沒綁過也沒給→ 回 need_name，前端再問一次
   */
  if (p.action === 'google') {
    var glk = LockService.getScriptLock();
    try { glk.waitLock(20000); } catch (err) { return out({ ok: false, error: 'busy' }, cb); }
    try {
      var g = verifyGoogle(p.idt);
      if (!g.ok) return out(g, cb);

      var hit = findByGoogle(ss, g.sub);
      if (hit.row > 0) {
        var tk = issueToken(hit.sh, hit.row, hit.name);
        hit.sh.getRange(hit.row, U_USED).setValue(nowText());
        markLogin(ss, hit);
        var st1 = daysOf(ss, hit.name);
        return out({ ok: true, created: false, name: hit.name, role: hit.role,
                     token: tk, days: st1.days, count: st1.count }, cb);
      }

      // 綁定既有帳號：要能證明那個帳號是他的，所以照樣驗一次密碼
      if (p.link) {
        var la = auth(ss, p.link, p.linkpin, false);
        if (!la.ok) return out({ ok: false, error: la.error }, cb);
        if (la.gid) return out({ ok: false, error: 'already_linked' }, cb);
        la.sh.getRange(la.row, U_GID).setValue(g.sub);
        var tk2 = issueToken(la.sh, la.row, la.name);
        markLogin(ss, la);
        var st2 = daysOf(ss, la.name);
        return out({ ok: true, created: false, linked: true, name: la.name, role: la.role,
                     token: tk2, days: st2.days, count: st2.count }, cb);
      }

      var nn2 = cleanName(p.name);
      if (!nn2) return out({ ok: false, error: 'need_name', suggest: g.suggest }, cb);
      if (findUser(ss, nn2).row > 0) return out({ ok: false, error: 'exists' }, cb);
      var sh2 = sheetOf(ss, SH_USERS, HEAD_USERS);
      var tok2 = newToken();
      sh2.appendRow([nn2, hashPin(nn2, tok2), nowText(), nowText(), 'user', 0, '', g.sub]);
      var made = { sh: sh2, row: sh2.getLastRow(), name: nn2, role: 'user', logins: 0 };
      markLogin(ss, made);
      return out({ ok: true, created: true, name: nn2, role: 'user',
                   token: tok2, days: [], count: [] }, cb);
    } finally {
      glk.releaseLock();
    }
  }

  // 清自己的算牌記錄。只要帳密對就能清自己的，不需要系統人員。
  if (p.action === 'clear_count') {
    var ca = auth(ss, p.name, p.pin, false);
    if (!ca.ok) return out(ca, cb);
    var lk3 = LockService.getScriptLock();
    try { lk3.waitLock(30000); } catch (err) { return out({ ok: false, error: 'busy' }, cb); }
    try {
      return out({ ok: true, removed: clearCount(ss, ca.name) }, cb);
    } finally {
      lk3.releaseLock();
    }
  }

  // 診斷用：時區錯了每日統計會整個切錯天，出問題時先看這個
  if (p.action === 'tz') {
    var now = new Date();
    return out({
      ok: true,
      sheetTz: ss.getSpreadsheetTimeZone(),
      scriptTz: Session.getScriptTimeZone(),
      iso: now.toISOString(),
      bySheet: Utilities.formatDate(now, ss.getSpreadsheetTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
      byTaipei: Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm')
    }, cb);
  }

  if (p.action === 'admin_rebuild') {
    var ar = requireAdmin(ss, p);
    if (!ar.ok) return out(ar, cb);
    var lk2 = LockService.getScriptLock();
    try { lk2.waitLock(30000); } catch (err) { return out({ ok: false, error: 'busy' }, cb); }
    try {
      return out({ ok: true, rows: rebuildDaily(ss) }, cb);
    } finally {
      lk2.releaseLock();
    }
  }

  if (p.action === 'admin_delete') {
    var ad = requireAdmin(ss, p);
    if (!ad.ok) return out(ad, cb);
    var dn = cleanName(p.target);
    if (!dn) return out({ ok: false, error: 'bad_input' }, cb);
    // 不能砍自己，避免後台把自己鎖在外面
    if (dn === ad.name) return out({ ok: false, error: 'self' }, cb);
    var t2 = findUser(ss, dn);
    if (t2.row < 0) return out({ ok: false, error: 'no_user' }, cb);
    t2.sh.deleteRow(t2.row);
    // 練習記錄刻意保留：帳號刪掉只是不能再登入，歷史資料還在試算表裡可查
    return out({ ok: true, name: dn }, cb);
  }

  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<body style="font-family:-apple-system,sans-serif;background:#0F1214;color:#EDEBE6;padding:28px">' +
    '<h2 style="margin:0 0 10px">戰勝21點 · 記錄後端運作中</h2>' +
    '<p style="color:#9AA0A2;line-height:1.7">這是資料端點，不是練習頁面。<br>' +
    '練習請開 <a style="color:#C6A34B" href="https://linpoyes.github.io/blackjack-trainer/">' +
    'linpoyes.github.io/blackjack-trainer</a></p></body>'
  );
}
