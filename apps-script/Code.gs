/**
 * 一周日常表 · Weekly Schedule — Google Apps Script 版本
 * ------------------------------------------------------------------
 * 后端职责：
 *   1. doGet()          → 以 Web App 形式返回 Index.html 界面
 *   2. loadState()      → 从 Google Sheet 读取日程与设置（数据持久化在云端）
 *   3. saveState(json)  → 把界面状态写回 Google Sheet（自动保存）
 *   4. exportToSheet()  → 生成格式化的「日程表 Export」工作表（导出能力）
 *
 * 数据存储：一个 Google Sheet，含三个工作表
 *   Data          —— 每条日程一行（id/date/start/end/title/category/attendees）
 *   Meta          —— 界面设置（title/start/end/lang/view）
 *   日程表 Export  —— 导出的成品表格（标题、表头样式、合计、筛选）
 *
 * 首次调用会自动在 Drive 中新建这个表格；也可以把已有表格的 ID 填到下面的
 * SHEET_ID 常量来复用它。
 */

// ------------------------- 配置 -------------------------

/** 可选：填入已有 Google Sheet 的 ID 以复用它；留空则自动新建。 */
var SHEET_ID = '';

var PROP_KEY = 'weekly-cal.sheet.id';
var SHEET_DATA = 'Data';
var SHEET_META = 'Meta';
var SHEET_EXPORT = '日程表 Export';

var DATA_HEADERS = ['id', 'date', 'start', 'end', 'title', 'category', 'attendees'];
var DATA_FORMATS = ['@', '@', '0', '0', '@', '@', '@']; // start/end 为分钟数，其余为文本
var META_KEYS = ['title', 'start', 'end', 'lang', 'view'];

// To-B 设计规范色值（与前端一致）
var BRAND = '#3643BA';
var BORDER = '#D9D9D9';
var ROW_ALT = '#F6F6F6';
var TEXT_TERTIARY = '#8C8C8C';

var L10N = {
  zh: {
    title: '一周日常表',
    generated: '导出时间',
    range: '日期范围',
    headers: ['日期', '星期', '开始', '结束', '时长', '主题', '类别', '与会人员'],
    weekdays: ['周日', '周一', '周二', '周三', '周四', '周五', '周六'],
    cats: { work: '工作 / 学习', sport: '运动', food: '用餐', rest: '休息', personal: '个人事务' },
    total: '合计',
    totalEvents: '共 {n} 条日程',
    empty: '（暂无日程：先在页面上按住拖拽框选时间段即可创建）',
    unit: [' 小时 ', ' 分钟'],
    font: 'Source Han Sans CN'
  },
  en: {
    title: 'Weekly Schedule',
    generated: 'Exported at',
    range: 'Date range',
    headers: ['Date', 'Weekday', 'Start', 'End', 'Duration', 'Title', 'Category', 'Attendees'],
    weekdays: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    cats: { work: 'Work / Study', sport: 'Sport', food: 'Meal', rest: 'Rest', personal: 'Personal' },
    total: 'Total',
    totalEvents: '{n} event(s)',
    empty: '(No events yet — press and drag a time range on the page to create one)',
    unit: ['h ', 'm'],
    font: 'SF Pro Text'
  }
};

// ------------------------- Web App 入口 -------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('一周日常表 · Weekly Calendar')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 可选：在编辑器里手动运行一次，提前创建表格并在执行日志中打印链接。 */
function bootstrap() {
  var ss = getSpreadsheet_();
  var msg = '表格已就绪: ' + ss.getUrl();
  Logger.log(msg);
  return msg;
}

// ------------------------- 状态读写 -------------------------

/**
 * 前端调用。返回 { json: <状态 JSON 字符串>, url: <表格链接> }
 */
function loadState() {
  var ss = getSpreadsheet_();
  var meta = readMeta_(ss);
  var state = {
    events: readEvents_(ss),
    title: meta.title || '',
    start: meta.start || '',
    end: meta.end || '',
    lang: meta.lang === 'en' ? 'en' : 'zh',
    view: meta.view || ''
  };
  return { json: JSON.stringify(state), url: ss.getUrl() };
}

/**
 * 前端调用。json 为界面完整状态的 JSON 字符串。
 */
function saveState(json) {
  var st = JSON.parse(json); // 解析失败会作为错误回传前端
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    throw new Error('表格正忙，请稍后重试 (lock timeout)');
  }
  try {
    var ss = getSpreadsheet_();
    writeEvents_(ss, st.events || []);
    writeMeta_(ss, st);
    return true;
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* ignore */ }
  }
}

// ------------------------- 导出到 Google Sheet -------------------------

/**
 * 生成 / 刷新格式化的导出工作表。
 * @param {string} lang 'zh' | 'en'
 * @return {{url: string, sheet: string, rows: number}}
 */
function exportToSheet(lang) {
  var L = L10N[lang === 'en' ? 'en' : 'zh'];
  var ss = getSpreadsheet_();
  var meta = readMeta_(ss);
  var events = readEvents_(ss).sort(function (a, b) {
    if (a.date === b.date) return a.start - b.start;
    return a.date < b.date ? -1 : 1;
  });

  var sh = ss.getSheetByName(SHEET_EXPORT);
  if (!sh) sh = ss.insertSheet(SHEET_EXPORT);
  // 旧筛选器必须先移除，否则 clear() 后仍会残留；任何异常都不应中断导出
  try {
    var oldFilter = sh.getFilter();
    if (oldFilter) oldFilter.remove();
  } catch (eFilter) { /* 无筛选器或已失效 */ }
  sh.clear();

  var cols = L.headers.length; // 8

  // 第 1 行：日程表标题（跟随页面上可编辑的标题）
  sh.getRange(1, 1).setValue((meta.title || '').trim() || L.title);
  sh.getRange(1, 1, 1, cols).merge()
    .setFontWeight('bold').setFontSize(16).setFontFamily(L.font);

  // 第 2 行：日期范围 + 导出时间
  var rangeText = (meta.start && meta.end) ? (meta.start + ' → ' + meta.end) : '—';
  var tz = Session.getScriptTimeZone() || 'Asia/Shanghai';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  sh.getRange(2, 1).setValue(L.range + ': ' + rangeText + '     ·     ' + L.generated + ': ' + stamp);
  sh.getRange(2, 1, 1, cols).merge()
    .setFontSize(10).setFontColor(TEXT_TERTIARY).setFontFamily(L.font);

  // 表头
  var headerRow = 4;
  sh.getRange(headerRow, 1, 1, cols).setValues([L.headers])
    .setBackground(BRAND).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(11).setFontFamily(L.font)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(headerRow, 28);

  // 数据行
  var rows = events.map(function (ev) {
    var d = parseDate_(ev.date);
    return [
      ev.date,
      d ? L.weekdays[d.getDay()] : '',
      fmtMin_(ev.start),
      fmtMin_(ev.end),
      fmtDur_(Math.max(0, num_(ev.end) - num_(ev.start)), L),
      text_(ev.title),
      L.cats[ev.cat] || text_(ev.cat),
      text_(ev.att)
    ];
  });

  var bodyStart = headerRow + 1;
  if (rows.length) {
    var body = sh.getRange(bodyStart, 1, rows.length, cols);
    body.setNumberFormats(rows.map(function () { return ['@', '@', '@', '@', '@', '@', '@', '@']; }))
      .setValues(rows).setFontSize(11).setFontFamily(L.font).setVerticalAlignment('middle');
    sh.getRange(bodyStart, 1, rows.length, 5).setHorizontalAlignment('center');
    sh.getRange(bodyStart, 6, rows.length, 3).setHorizontalAlignment('left');
    for (var i = 0; i < rows.length; i += 2) {
      sh.getRange(bodyStart + i, 1, 1, cols).setBackground(ROW_ALT);
    }
  } else {
    sh.getRange(bodyStart, 1).setValue(L.empty);
    sh.getRange(bodyStart, 1, 1, cols).merge()
      .setFontColor(TEXT_TERTIARY).setFontSize(11).setFontFamily(L.font)
      .setHorizontalAlignment('center');
    sh.setRowHeight(bodyStart, 28);
  }

  // 合计行
  var totalMinutes = events.reduce(function (sum, ev) {
    return sum + Math.max(0, num_(ev.end) - num_(ev.start));
  }, 0);
  var totalRow = bodyStart + Math.max(rows.length, 1);
  sh.getRange(totalRow, 1).setValue(L.total);
  sh.getRange(totalRow, 5).setValue(fmtDur_(totalMinutes, L)).setHorizontalAlignment('center');
  sh.getRange(totalRow, 6).setValue(L.totalEvents.replace('{n}', String(events.length)));
  sh.getRange(totalRow, 1, 1, cols)
    .setFontWeight('bold').setFontSize(11).setFontFamily(L.font);

  // 表格区域边框
  var tableLast = totalRow - 1;
  if (tableLast >= headerRow) {
    sh.getRange(headerRow, 1, tableLast - headerRow + 1, cols)
      .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
  sh.getRange(totalRow, 1, 1, cols)
    .setBorder(false, false, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // 列宽 / 冻结表头 / 筛选
  var widths = [96, 64, 64, 64, 96, 240, 112, 200];
  for (var c = 0; c < cols; c++) sh.setColumnWidth(c + 1, widths[c]);
  sh.setFrozenRows(headerRow);
  if (rows.length) {
    try { sh.getRange(headerRow, 1, rows.length + 1, cols).createFilter(); } catch (eFilter2) { /* ignore */ }
  }

  ss.setActiveSheet(sh);
  return { url: ss.getUrl(), sheet: SHEET_EXPORT, rows: events.length };
}

// ------------------------- 表格结构 -------------------------

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = SHEET_ID || props.getProperty(PROP_KEY);
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('一周日常表 · Weekly Schedule');
    props.setProperty(PROP_KEY, ss.getId());
  }
  ensureSheets_(ss);
  return ss;
}

function ensureSheets_(ss) {
  var data = ss.getSheetByName(SHEET_DATA);
  if (!data) data = ss.insertSheet(SHEET_DATA);
  if (data.getLastRow() === 0) {
    data.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS])
      .setFontWeight('bold').setBackground(ROW_ALT);
    data.setFrozenRows(1);
  }

  var meta = ss.getSheetByName(SHEET_META);
  if (!meta) meta = ss.insertSheet(SHEET_META);
  if (meta.getLastRow() === 0) {
    meta.getRange(1, 1, 1, 2).setValues([['key', 'value']])
      .setFontWeight('bold').setBackground(ROW_ALT);
    meta.setFrozenRows(1);
    meta.setColumnWidth(1, 120);
    meta.setColumnWidth(2, 320);
  }

  // 删除自动新建的空白默认工作表（仅在确有其它工作表时）
  var sheets = ss.getSheets();
  if (sheets.length > 2) {
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name !== SHEET_DATA && name !== SHEET_META && name !== SHEET_EXPORT &&
          sheets[i].getLastRow() === 0) {
        try { ss.deleteSheet(sheets[i]); } catch (e) { /* ignore */ }
        break;
      }
    }
  }
}

function readEvents_(ss) {
  var sh = ss.getSheetByName(SHEET_DATA);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, DATA_HEADERS.length).getValues();
  var out = [];
  values.forEach(function (r) {
    if (!text_(r[0]).length) return;
    out.push({
      id: text_(r[0]),
      date: dateText_(r[1]),
      start: num_(r[2]),
      end: num_(r[3]),
      title: text_(r[4]),
      cat: text_(r[5]),
      att: text_(r[6])
    });
  });
  return out;
}

function writeEvents_(ss, events) {
  var sh = ss.getSheetByName(SHEET_DATA);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, DATA_HEADERS.length).clearContent();
  if (!events.length) return;
  var rows = events.map(function (e) {
    return [
      text_(e.id), dateText_(e.date), num_(e.start), num_(e.end),
      text_(e.title), text_(e.cat), text_(e.att)
    ];
  });
  var range = sh.getRange(2, 1, rows.length, DATA_HEADERS.length);
  range.setNumberFormats(rows.map(function () { return DATA_FORMATS; }));
  range.setValues(rows);
}

function readMeta_(ss) {
  var sh = ss.getSheetByName(SHEET_META);
  var last = sh.getLastRow();
  var out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    var k = text_(r[0]);
    if (!k) return;
    out[k] = r[1] instanceof Date ? dateText_(r[1]) : text_(r[1]);
  });
  return out;
}

function writeMeta_(ss, st) {
  var sh = ss.getSheetByName(SHEET_META);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  var rows = META_KEYS.map(function (k) { return [k, text_(st[k])]; });
  var range = sh.getRange(2, 1, rows.length, 2);
  range.setNumberFormat('@');
  range.setValues(rows);
}

// ------------------------- 工具函数 -------------------------

function num_(v) {
  var n = parseInt(v, 10);
  return isNaN(n) ? 0 : n;
}

function text_(v) {
  return (v === null || v === undefined) ? '' : String(v);
}

/** 单元格可能被 Sheets 解析成 Date，统一还原为 YYYY-MM-DD 文本。 */
function dateText_(v) {
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone() || 'Asia/Shanghai';
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return text_(v);
}

function parseDate_(str) {
  var p = String(str || '').split('-');
  if (p.length !== 3) return null;
  var d = new Date(+p[0], (+p[1]) - 1, +p[2]);
  return isNaN(d.getTime()) ? null : d;
}

/** 分钟数 → HH:MM */
function fmtMin_(m) {
  m = num_(m);
  var h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
}

/** 分钟数 → 时长文案，跟随语言 */
function fmtDur_(min, L) {
  min = num_(min);
  var h = Math.floor(min / 60), m = min % 60;
  var u = L.unit; // [小时单位, 分钟单位]
  if (h && m) return h + u[0] + m + u[1];
  if (h) return h + u[0].replace(/\s+$/, '');
  return m + u[1];
}
