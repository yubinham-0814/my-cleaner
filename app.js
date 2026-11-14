/* app.js — 화면
 *
 * 파일 넣기 → 줄 수 · 칸 수 → 처방 보여 주기 → 규칙대로 정리 → 결과 · 내려받기 → 정답과 대조
 * 파일 내용은 이 브라우저 밖으로 나가지 않는다. 이 파일이 서버에 요청하는 것은 둘뿐이다.
 *   1. 화면을 열 때 처방 목록 rules/index.json
 *   2. 엑셀을 넣었을 때, 그 이름이 처방 목록에 있으면 그 처방 파일 rules/<이름>.json 하나
 *      (목록에 없는 이름이면 아무것도 요청하지 않는다)
 */
(function () {
  'use strict';

  var C = window.Cleaner;
  var state = { ruleIndex: [], ruleIndexError: null, file: null, wb: null, rulesSet: null, table: null, result: null, idem: null };

  function $(id) { return document.getElementById(id); }

  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === 'class') e.className = attrs[k];
      else if (k === 'text') e.textContent = attrs[k];
      else e.setAttribute(k, attrs[k]);
    }
    (children || []).forEach(function (ch) { e.appendChild(typeof ch === 'string' ? document.createTextNode(ch) : ch); });
    return e;
  }

  function showBox(id, on) { $(id).hidden = !on; }
  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }
  function fmtNum(n) { return Number(n).toLocaleString('ko-KR'); }

  function errorText(err) {
    if (err && err.name === 'CleanError') return err.message;
    return '처리하는 중에 문제가 생겼습니다: ' + (err && err.message ? err.message : String(err));
  }

  // ------------------------------------------------------------ 처방 목록 읽기
  // 화면을 열 때 처방 목록(이름들)만 읽는다. 처방 파일 자체는 엑셀을 넣었을 때 읽는다.
  function loadRuleIndex() {
    if (!window.fetch || location.protocol === 'file:') {
      state.ruleIndexError = '이 화면을 파일로 직접 열어서 처방 목록을 읽지 못했습니다. 엑셀을 넣은 뒤 처방 파일을 직접 골라 주세요.';
      $('rules-status').textContent = state.ruleIndexError;
      return Promise.resolve();
    }
    return fetch('rules/index.json', { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('rules/index.json ' + r.status); return r.json(); })
      .then(function (idx) {
        state.ruleIndex = ((idx && idx.files) || []).map(function (name) { return { name: name }; });
        // 잘 읽었으면 아무것도 띄우지 않는다. (엑셀을 넣기 전에 «읽었습니다»가 보이면 헷갈린다)
        $('rules-status').textContent = '';
      })
      .catch(function (e) {
        state.ruleIndexError = '처방 목록(rules/index.json)을 읽지 못했습니다: ' + (e.message || e);
        $('rules-status').textContent = state.ruleIndexError;
      });
  }
  var ruleIndexReady = loadRuleIndex();

  // 엑셀 이름이 처방 목록에 있을 때만 그 처방 파일 하나를 요청한다.
  // 목록에 없는 이름이면 아무것도 요청하지 않으므로, 넣은 파일 이름이 밖으로 나가지 않는다.
  function findAndLoadRules(file) {
    if (state.file !== file) return;
    showBox('step-rules', true);
    var found = C.findRules(file.name, state.ruleIndex);
    if (!found) { showRulesMissing(); return; }
    return fetch('rules/' + encodeURIComponent(found.name), { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function (rules) {
        if (state.file === file) useRules({ name: found.name, rules: rules, source: 'rules/' + found.name });
      })
      .catch(function (e) {
        if (state.file === file) showRulesMissing('처방 파일 rules/' + found.name + '을 읽지 못했습니다 (' + (e.message || e) + ').');
      });
  }

  // ------------------------------------------------------------ 1. 파일 넣기

  function resetAfterFile() {
    state.rulesSet = null; state.table = null; state.result = null; state.idem = null;
    ['step-rules', 'step-result', 'step-answer', 'rules-missing', 'rules-errors', 'rules-view', 'run-error', 'result-view', 'answer-view', 'answer-error'].forEach(function (id) { showBox(id, false); });
    $('answer-input').value = '';
  }

  function handleFile(file) {
    if (!file) return;
    resetAfterFile();
    state.file = file;
    var info = $('file-info');
    clear(info);
    info.appendChild(el('p', { text: '읽는 중: ' + file.name }));
    showBox('file-info', true);
    file.arrayBuffer().then(function (buf) {
      state.wb = C.workbookFrom(new Uint8Array(buf));
      renderFileInfo();
    }).catch(function (err) {
      clear(info);
      info.appendChild(el('div', { class: 'notice error', text: '엑셀 파일로 읽지 못했습니다. (' + (err.message || err) + ')' }));
      state.wb = null;
    }).then(function () {
      if (!state.wb || state.file !== file) return;
      return ruleIndexReady.then(function () { return findAndLoadRules(file); });
    });
  }

  function renderFileInfo() {
    var info = $('file-info');
    clear(info);
    var sizes = C.sheetSizes(state.wb);
    info.appendChild(el('p', {}, [el('strong', { text: state.file.name })]));
    var stats = el('div', { class: 'stats' });
    sizes.forEach(function (s) {
      stats.appendChild(stat('시트 «' + s.name + '»', fmtNum(s.rows) + '줄', '칸 ' + s.cols + '개 · 머리글 제외' + (s.skippedEmpty ? ' · 빈 줄 ' + s.skippedEmpty + '개 건너뜀' : '')));
    });
    info.appendChild(stats);
  }

  function stat(label, value, sub) {
    return el('div', { class: 'stat' }, [
      el('div', { class: 'label', text: label }),
      el('div', { class: 'value', text: value }),
      sub ? el('div', { class: 'sub', text: sub }) : el('span')
    ]);
  }

  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (t) {
    drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (t) {
    drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('over'); });
  });
  drop.addEventListener('drop', function (e) {
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    handleFile(f);
  });
  // 화면 다른 곳에 떨어뜨려도 브라우저가 파일을 열어 버리지 않게 한다
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) { e.preventDefault(); });
  $('file-input').addEventListener('change', function (e) { handleFile(e.target.files[0]); e.target.value = ''; });

  // ------------------------------------------------------------ 2. 처방

  function showRulesMissing(reason) {
    showBox('rules-missing', true);
    showBox('rules-view', false);
    var d = reason || ('찾은 이름: rules/' + C.baseName(state.file.name) + '.json');
    if (state.ruleIndexError) d += ' · ' + state.ruleIndexError;
    $('rules-missing-detail').textContent = d + ' · 처방 파일을 직접 골라 올릴 수 있습니다.';
  }

  $('rules-input').addEventListener('change', function (e) {
    var f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    f.text().then(function (txt) {
      var rules;
      try { rules = JSON.parse(txt.replace(/^﻿/, '')); }
      catch (err) { showRulesErrors(['JSON으로 읽지 못했습니다: ' + err.message]); return; }
      useRules({ name: f.name, rules: rules, source: '직접 고른 파일 ' + f.name });
    });
  });

  function showRulesErrors(list) {
    var box = $('rules-errors');
    box.textContent = '이 처방 파일로는 정리할 수 없습니다.\n- ' + list.join('\n- ');
    showBox('rules-errors', true);
  }

  function useRules(set) {
    showBox('rules-errors', false);
    showBox('step-result', false);
    showBox('step-answer', false);
    var problems = C.validateRules(set.rules);
    if (problems.length) { showRulesErrors(problems); showBox('rules-view', false); return; }
    var table;
    try { table = C.readTable(state.wb, set.rules); }
    catch (err) { showRulesErrors([errorText(err)]); return; }
    var trouble = [];
    if (table.missingColumns.length) trouble.push('처방이 쓰는 칸이 이 파일에 없습니다: ' + table.missingColumns.join(', '));
    if (table.duplicateHeaders.length) trouble.push('같은 이름의 칸이 두 개 있습니다: ' + table.duplicateHeaders.join(', '));
    if (trouble.length) { showRulesErrors(trouble); showBox('rules-view', false); return; }

    state.rulesSet = set;
    state.table = table;
    showBox('rules-missing', false);
    var r = set.rules;
    var meta = '처방 파일: ' + set.source + ' · 정한 사람: ' + (r.owner || '-') + ' · 고친 날: ' + (r.updated || '-') +
      ' · 시트: ' + table.sheetName + ' · ' + fmtNum(table.rows.length) + '줄';
    if (table.extraColumns.length) meta += ' · 처방에 없는 칸(그대로 둠): ' + table.extraColumns.join(', ');
    $('rules-meta').textContent = meta;
    var list = $('rules-list');
    clear(list);
    r.rules.forEach(function (rule) {
      list.appendChild(el('li', {}, [
        el('div', {}, [el('span', { class: 'what', text: rule.what || rule.id }), el('span', { class: 'id', text: rule.id })]),
        el('div', { class: 'why', text: '이유: ' + rule.why + (rule.owner ? ' · ' + rule.owner : '') })
      ]));
    });
    var fc = r.flag_column;
    list.appendChild(el('li', {}, [
      el('div', { class: 'what', text: '확인이 필요한 줄은 «' + fc.name + '» 칸에 «' + (fc.message || '') + '»라고 적습니다' }),
      el('div', { class: 'why', text: fc.note || '' })
    ]));
    showBox('rules-view', true);
  }

  // ------------------------------------------------------------ 3. 정리

  $('run').addEventListener('click', function () {
    showBox('step-result', true);
    showBox('run-error', false);
    showBox('result-view', false);
    var rules = state.rulesSet.rules;
    try {
      state.result = C.clean(state.table, rules, { today: C.todayIso() });
      state.idem = C.checkIdempotent(state.result, rules);
    } catch (err) {
      $('run-error').textContent = errorText(err);
      showBox('run-error', true);
      return;
    }
    renderResult();
    showBox('step-answer', true);
    showBox('answer-view', false);
    showBox('answer-error', false);
  });

  function renderResult() {
    var res = state.result;
    var stats = $('stats');
    clear(stats);
    stats.appendChild(stat('정리 전', fmtNum(res.before) + '줄'));
    stats.appendChild(stat('정리 후', fmtNum(res.after) + '줄'));
    stats.appendChild(stat('지운 행', fmtNum(res.removed.length)));
    stats.appendChild(stat('고친 값', fmtNum(res.changedCellCount) + '칸', fmtNum(res.changedRowCount) + '줄에서'));
    stats.appendChild(stat('표시한 행', fmtNum(res.flagged.length)));
    stats.appendChild(stat('기준 날짜', res.today, '«오늘 이후» 판단에 씀'));

    var idem = $('idem');
    clear(idem);
    if (state.idem.same) {
      idem.className = 'notice ok';
      idem.appendChild(el('p', { text: '✓ 두 번 돌려도 같은 결과: 정리본을 다시 넣고 한 번 더 정리해도 바뀌는 것이 없습니다.' }));
    } else {
      idem.className = 'notice error';
      idem.appendChild(el('p', { text: '✗ 두 번 돌리면 결과가 달라집니다. 처방을 확인해야 합니다.' }));
      state.idem.diffs.slice(0, 10).forEach(function (d) { idem.appendChild(el('p', { class: 'small', text: d })); });
    }

    var key = state.rulesSet.rules.key;
    fillTable('tbl-flagged', ['원래 행', key, res.flagName, '이유'], res.flagged.map(function (f) {
      return [f.rowNum, f.key, f.text, f.whys.join(' / ')];
    }), { wrap: [3] });
    $('count-flagged').textContent = '(' + res.flagged.length + ')';

    fillTable('tbl-changes', ['원래 행', key, '칸', '바꾸기 전', '바꾼 후', '이유'], res.changes.map(function (ch) {
      return [ch.rowNum, ch.key, ch.column, C.show(ch.before), C.show(ch.after, res.formats[ch.column]), ch.why];
    }), { wrap: [5], mark: [3, 4] });
    $('count-changes').textContent = '(' + res.changes.length + '건 · ' + res.changedCellCount + '칸)';

    fillTable('tbl-removed', ['원래 행', key, '남긴 행', '이유'], res.removed.map(function (r) {
      return [r.rowNum, r.key, r.keptRowNum, r.why];
    }), { wrap: [3] });
    $('count-removed').textContent = '(' + res.removed.length + ')';

    showBox('warn-box', res.warnings.length > 0);
    fillTable('tbl-warnings', ['원래 행', key, '칸', '값', '까닭'], res.warnings.map(function (w) {
      return [w.rowNum, w.key, w.column, w.value, w.note];
    }), { wrap: [4] });
    $('count-warnings').textContent = '(' + res.warnings.length + ')';

    showBox('result-view', true);
  }

  function fillTable(id, header, rows, opt) {
    opt = opt || {};
    var t = $(id);
    clear(t);
    var thead = el('thead', {}, [el('tr', {}, header.map(function (h) { return el('th', { text: h }); }))]);
    var tbody = el('tbody');
    if (!rows.length) {
      tbody.appendChild(el('tr', {}, [el('td', { colspan: String(header.length), text: '없음' })]));
    }
    rows.forEach(function (r) {
      tbody.appendChild(el('tr', {}, r.map(function (v, i) {
        var cls = [];
        if (typeof v === 'number') cls.push('num');
        if ((opt.wrap || []).indexOf(i) >= 0) cls.push('wrap');
        var td = el('td', cls.length ? { class: cls.join(' ') } : null);
        var text = v === null || v === undefined ? '' : String(v);
        if ((opt.mark || []).indexOf(i) >= 0 && text !== '') td.appendChild(el('span', { class: 'mark', text: text.replace(/ /g, '·') }));
        else td.textContent = text;
        return td;
      })));
    });
    t.appendChild(thead);
    t.appendChild(tbody);
  }

  // ------------------------------------------------------------ 내려받기 (이 컴퓨터에 바로 저장)

  function outName(suffix) { return C.baseName(state.file.name) + '_' + suffix + '.xlsx'; }

  $('dl-clean').addEventListener('click', function () {
    XLSX.writeFile(C.buildCleanWorkbook(state.result), outName('정리본'), { compression: true });
  });
  $('dl-log').addEventListener('click', function () {
    var wb = C.buildLogWorkbook(state.result, state.rulesSet.rules, {
      fileName: state.file.name, rulesName: state.rulesSet.source, idempotent: state.idem ? state.idem.same : null
    });
    XLSX.writeFile(wb, outName('바꾼기록'), { compression: true });
  });

  // ------------------------------------------------------------ 4. 정답과 대조

  function readAnswer(file) {
    var isCsv = /\.csv$/i.test(file.name);
    return file.arrayBuffer().then(function (buf) {
      if (!isCsv) return C.workbookFrom(new Uint8Array(buf));
      var text = new TextDecoder('utf-8').decode(buf);
      if (text.indexOf('�') >= 0) {
        try { text = new TextDecoder('euc-kr').decode(buf); } catch (e) { /* utf-8 그대로 */ }
      }
      text = text.replace(/^﻿/, '');
      return XLSX.read(text, { type: 'string', raw: true });
    });
  }

  $('answer-input').addEventListener('change', function (e) {
    var f = e.target.files[0];
    if (!f || !state.result) return;
    showBox('answer-error', false);
    showBox('answer-view', false);
    readAnswer(f).then(function (wb) {
      var cmp = C.compareWithAnswer(state.result, wb, state.rulesSet.rules);
      renderCompare(f.name, cmp);
    }).catch(function (err) {
      $('answer-error').textContent = '정답 파일을 읽지 못했습니다: ' + (err.message || err);
      showBox('answer-error', true);
    });
  });

  function renderCompare(name, cmp) {
    var s = $('answer-stats');
    clear(s);
    s.appendChild(stat('정답 행', fmtNum(cmp.answerRows) + '줄', name));
    s.appendChild(stat('정리본 행', fmtNum(cmp.ourRows) + '줄'));
    s.appendChild(stat('일치한 행', fmtNum(cmp.matchedRows) + '줄', '맞대 본 ' + fmtNum(cmp.compared) + '줄 중'));
    s.appendChild(stat('어긋난 칸', fmtNum(cmp.mismatches.length) + '칸'));
    var notes = [];
    notes.push(cmp.keyUsed ? '«' + cmp.keyUsed + '»로 줄을 맞대 봤습니다.' : '«' + state.rulesSet.rules.key + '» 칸이 정답에 없어 줄 순서대로 맞대 봤습니다.');
    notes.push('비교한 칸: ' + cmp.commonColumns.join(', '));
    if (cmp.onlyOurCols.length) notes.push('정리본에만 있는 칸: ' + cmp.onlyOurCols.join(', '));
    if (cmp.onlyAnswerCols.length) notes.push('정답에만 있는 칸: ' + cmp.onlyAnswerCols.join(', '));
    if (cmp.onlyAnswer.length) notes.push('정답에만 있는 줄 ' + cmp.onlyAnswer.length + ': ' + cmp.onlyAnswer.slice(0, 20).map(function (r) { return r.key; }).join(', '));
    if (cmp.onlyOurs.length) notes.push('정리본에만 있는 줄 ' + cmp.onlyOurs.length + ': ' + cmp.onlyOurs.slice(0, 20).map(function (r) { return r.key; }).join(', '));
    if (cmp.dupAnswerKeys.length) notes.push('정답에서 두 번 이상 나온 ' + cmp.keyUsed + ': ' + cmp.dupAnswerKeys.slice(0, 20).join(', '));
    notes.push('숫자는 쉼표를 빼고, 날짜는 YYYY-MM-DD로 맞춰 비교했습니다. 글자는 공백까지 그대로 비교했습니다.');
    $('answer-notes').textContent = notes.join(' · ');
    fillTable('tbl-mismatch', ['줄', '칸', '정리본', '정답'], cmp.mismatches.map(function (m) {
      return [m.key, m.column, m.ours, m.answer];
    }), { mark: [2, 3] });
    $('count-mismatch').textContent = '(' + cmp.mismatches.length + ')';
    showBox('answer-view', true);
  }
})();
