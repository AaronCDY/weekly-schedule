#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build Apps Script Index.html from the standalone weekly-schedule.html.

Keeps every existing feature (drag-create, editable title, date range,
attendees, visible-hours presets, zh/en, print) and swaps the storage layer
from localStorage to google.script.run -> Google Sheet.
"""
import os
import re
import sys

# 默认按仓库布局：standalone 页面在上一级目录的 index.html。
# 也可用参数覆盖：python3 build-index.py <src.html> [dst.html]
BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, '..', 'index.html')
DST = os.path.join(BASE, 'Index.html')
if len(sys.argv) > 1:
    SRC = sys.argv[1]
if len(sys.argv) > 2:
    DST = sys.argv[2]

src = open(SRC, encoding='utf-8').read()

# ---- split document into style + body -------------------------------------
m = re.search(r'<style>(.*?)</style>', src, re.S)
assert m, 'style block not found'
css = m.group(1)

m = re.search(r'<body>(.*?)</body>', src, re.S)
assert m, 'body not found'
body = m.group(1)

FONT_IMPORT = ("@import url('https://fonts.googleapis.com/css2?family="
               "Noto+Sans+SC:wght@400;500;700&display=swap');\n")

EXTRA_CSS = """
/* ---- Apps Script 版本新增（沿用 To-B token；无基线值处标注为实例值） ---- */
.sheet-link {
  font-size: 14px; line-height: 20px; color: var(--tob-brand);
  text-decoration: none; border-bottom: 1px solid var(--tob-brand);
}
.sheet-link:hover { color: var(--tob-text); border-bottom-color: var(--tob-text); }
.sheet-link[hidden] { display: none; }
/* 禁用态无规范基线值，采用实例值 opacity 0.5 */
.btn:disabled { opacity: 0.5; cursor: not-allowed; }
"""

out = '<style>\n' + FONT_IMPORT + css + EXTRA_CSS + '</style>\n' + body

n0 = len(out)


def rep(old, new, label):
    global out
    c = out.count(old)
    if c != 1:
        sys.exit('FAIL [%s]: expected 1 occurrence, found %d' % (label, c))
    out = out.replace(old, new, 1)
    print('ok   %s' % label)


# ---- R1: state load -> hydrate bridge -------------------------------------
rep("""  var state = { events: [] };
  try {
    var raw = localStorage.getItem(STORE_KEY);
    if (raw) state = JSON.parse(raw);
  } catch (e) { /* ignore */ }
  if (state.lang === 'en' || state.lang === 'zh') LANG = state.lang;
""",
    """  // 运行环境：Apps Script Web App 中数据存于 Google Sheet；
  // 直接用浏览器打开该文件时自动退化为 localStorage。
  var IS_GAS = typeof google !== 'undefined' && !!(google.script && google.script.run);
  var pendingSave = false;
  var state = { events: [] };

  function hydrate(obj) {
    if (obj && typeof obj === 'object') state = obj;
    if (!Array.isArray(state.events)) state.events = [];
    if (state.lang === 'en' || state.lang === 'zh') LANG = state.lang;
    normalize();
    return state;
  }
  function hydrateLocal() {
    try {
      var raw = localStorage.getItem(STORE_KEY);
      if (raw) return hydrate(JSON.parse(raw));
    } catch (e) { /* 沙箱中 localStorage 可能不可用 */ }
    normalize();
    return state;
  }
""", 'R1 state bridge')

# ---- R2: normalization -> normalize() -------------------------------------
rep("""  // seed defaults (current week) when unset; normalize any stored range
  if (!state.start || !state.end) {
    var w0 = weekDates();
    state.start = dateKey(w0[0]); state.end = dateKey(w0[6]);
  } else {
    var nr = normRange(state.start, state.end); state.start = nr.start; state.end = nr.end;
  }
  if (typeof state.title !== 'string' || !state.title.trim()) state.title = DEFAULT_TITLE[LANG];
  if (!VIEW_PRESETS.some(function (p) { return p.s + '-' + p.e === state.view; })) state.view = DEFAULT_VIEW;
""",
    """  // seed defaults (current week) when unset; normalize any stored range
  function normalize() {
    if (!state.start || !state.end) {
      var w0 = weekDates();
      state.start = dateKey(w0[0]); state.end = dateKey(w0[6]);
    } else {
      var nr = normRange(state.start, state.end); state.start = nr.start; state.end = nr.end;
    }
    if (typeof state.title !== 'string' || !state.title.trim()) state.title = DEFAULT_TITLE[LANG];
    if (!VIEW_PRESETS.some(function (p) { return p.s + '-' + p.e === state.view; })) state.view = DEFAULT_VIEW;
  }
""", 'R2 normalize()')

# ---- R3: persist -> Sheets + localStorage cache ---------------------------
rep("""  function persist() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify(state));
      flash(t('saved'));
    } catch (e) { flash(t('saveFail')); }
  }
""",
    """  var saveTimer = null;
  function persist() {
    var json = JSON.stringify(state);
    try { localStorage.setItem(STORE_KEY, json); } catch (e) { /* 离线缓存，失败可忽略 */ }
    if (!IS_GAS) { flash(t('savedLocal')); return; }   // 浏览器直接打开时的退化路径
    flash(t('saving'));
    pendingSave = true;
    if (saveTimer) clearTimeout(saveTimer);
    // 合并连续修改（拖拽 / 连续编辑），减少对表格的写入次数
    saveTimer = setTimeout(function () {
      saveTimer = null;
      google.script.run
        .withSuccessHandler(function () { pendingSave = false; flash(t('saved')); })
        .withFailureHandler(function (err) { pendingSave = false; flash(errText(t('saveFail'), err)); })
        .saveState(json);
    }, 600);
  }
  function errText(prefix, err) {
    var m = err && err.message ? err.message : (err ? String(err) : '');
    return m ? prefix + ' · ' + m : prefix;
  }
""", 'R3 persist()')

# ---- R4: print button -> popup print inside Apps Script -------------------
rep("""  document.getElementById('btn-print').addEventListener('click', function () { window.print(); });
""",
    """  // 在 Apps Script 的 iframe 沙箱里 window.print() 会打印宿主页面，
  // 因此改为把「当前显示的内容」复制到新窗口再打印，保证所见即所得。
  function printView() {
    if (!IS_GAS) { window.print(); return; }
    var cssText = Array.prototype.map.call(document.querySelectorAll('style'), function (s) {
      return s.innerHTML;
    }).join('\\n');
    var bodyHtml = ['.hero', '.cal-card'].map(function (sel) {
      var el = document.querySelector(sel);
      return el ? el.outerHTML : '';
    }).join('\\n');
    var w = window.open('', '_blank');
    if (!w) { window.print(); return; }
    w.document.open();
    w.document.write('<!DOCTYPE html><html lang="' + LANG + '"><head><meta charset="utf-8">' +
      '<title>' + document.getElementById('sched-title').textContent + '</title>' +
      '<style>' + cssText +
      '@page { size: A4 landscape; margin: 12mm; }' +
      'body { background: #fff; }' +
      '.page { max-width: none; padding: 0; }' +
      '.cal-scroll { max-height: none !important; overflow: visible !important; }' +
      '.toolbar, .overlay, .lang-switch { display: none !important; }' +
      '</style></head><body><div class="page">' + bodyHtml + '</div></body></html>');
    w.document.close();
    w.focus();
    setTimeout(function () { w.print(); }, 400);
  }
  document.getElementById('btn-print').addEventListener('click', printView);

  // ---------- 导出到 Google Sheet ----------
  var sheetLink = document.getElementById('sheet-link');
  function setSheetUrl(url) {
    if (!url) { sheetLink.hidden = true; return; }
    sheetLink.href = url;
    sheetLink.hidden = false;
  }
  document.getElementById('btn-export').addEventListener('click', function () {
    if (!IS_GAS) { flash(t('exportOnlyGas')); return; }
    var btn = this;
    btn.disabled = true;
    flash(t('exporting'));
    google.script.run
      .withSuccessHandler(function (res) {
        btn.disabled = false;
        if (res && res.url) setSheetUrl(res.url);
        flash(tpl('exported', { n: res ? res.rows : 0, sheet: res ? res.sheet : '' }));
      })
      .withFailureHandler(function (err) {
        btn.disabled = false;
        flash(errText(t('exportFail'), err));
      })
      .exportToSheet(LANG);
  });
""", 'R4 print + export')

# ---- R5: toolbar export button --------------------------------------------
rep("""      <button class="btn btn-primary" id="btn-print" type="button" data-i18n="btnPrint">打印 / 导出 PDF</button>
""",
    """      <button class="btn btn-ghost" id="btn-export" type="button" data-i18n="btnExport">导出到 Google Sheet</button>
      <button class="btn btn-primary" id="btn-print" type="button" data-i18n="btnPrint">打印 / 导出 PDF</button>
""", 'R5 export button')

# ---- R6: footer sheet link -------------------------------------------------
rep("""    <span id="save-state" data-i18n="saved">已自动保存</span>
""",
    """    <a class="sheet-link" id="sheet-link" href="#" target="_blank" rel="noopener" hidden data-i18n="sheetLink">打开表格</a>
    <span id="save-state" data-i18n="saved">已自动保存</span>
""", 'R6 footer sheet link')

# ---- R7: i18n zh -----------------------------------------------------------
rep("""      saved: '已自动保存', saveFail: '保存失败（浏览器存储不可用）', cleared: '已清空 {n} 条日程',
""",
    """      saved: '已自动保存到 Google Sheet', saveFail: '保存失败', cleared: '已清空 {n} 条日程',
      saving: '正在保存到 Google Sheet…', loaded: '已从 Google Sheet 载入',
      loadFail: '载入失败（当前显示本地缓存）', btnExport: '导出到 Google Sheet',
      exporting: '正在导出到 Google Sheet…', exported: '已导出 {n} 条日程到工作表「{sheet}」',
      exportFail: '导出失败', exportOnlyGas: '导出功能仅在 Apps Script 网页版中可用',
      sheetLink: '打开表格', savedLocal: '已自动保存（浏览器本地）',
""", 'R7 i18n zh')

# ---- R8: i18n en -----------------------------------------------------------
rep("""      saved: 'Saved', saveFail: 'Save failed (browser storage unavailable)', cleared: 'Cleared {n} event(s)',
""",
    """      saved: 'Saved to Google Sheet', saveFail: 'Save failed', cleared: 'Cleared {n} event(s)',
      saving: 'Saving to Google Sheet…', loaded: 'Loaded from Google Sheet',
      loadFail: 'Load failed (showing local cache)', btnExport: 'Export to Google Sheet',
      exporting: 'Exporting to Google Sheet…', exported: 'Exported {n} event(s) to sheet "{sheet}"',
      exportFail: 'Export failed', exportOnlyGas: 'Export is only available in the Apps Script web app',
      sheetLink: 'Open sheet', savedLocal: 'Saved locally',
""", 'R8 i18n en')

# ---- R9: init -> async remote load ----------------------------------------
rep("""  // ---------- init ----------
  setView(state.view, false);
  titleEl.textContent = state.title;
  syncRangeInputs();
  syncLangBtns();
  applyI18n();
  applyRange();
  scrollToNow();
  setInterval(renderNowLine, 60000);
})();
""",
    """  // ---------- init ----------
  var nowTimer = null;
  function initApp() {
    setView(state.view, false);
    titleEl.textContent = state.title;
    syncRangeInputs();
    syncLangBtns();
    applyI18n();
    applyRange();
    scrollToNow();
    if (!nowTimer) nowTimer = setInterval(renderNowLine, 60000);
  }

  // 先用本地缓存即时渲染（避免 iframe 白屏），再用云端数据覆盖。
  hydrateLocal();
  initApp();

  if (IS_GAS) {
    google.script.run
      .withSuccessHandler(function (res) {
        if (!res) return;
        if (res.url) setSheetUrl(res.url);
        if (!res.json || pendingSave) return;   // 用户已开始编辑时不覆盖
        var remote;
        try { remote = JSON.parse(res.json); } catch (e) { return; }
        hydrate(remote);
        initApp();
        flash(t('loaded'));
      })
      .withFailureHandler(function (err) { flash(errText(t('loadFail'), err)); })
      .loadState();
  }
})();
""", 'R9 async init')

open(DST, 'w', encoding='utf-8').write(out)
print('\nwrote %s (%d -> %d bytes)' % (DST, n0, len(out)))
