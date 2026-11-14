/* clean.js — 정리 엔진
 *
 * rules/*.json 의 처방을 읽어 그대로 실행한다.
 * - 처방(무엇을 어떻게 바꿀지)은 여기 적지 않는다. 여기에는 처방을 실행하는 방법만 있다.
 * - 파일 내용을 어디로도 보내지 않는다. 모든 계산은 브라우저 안에서 한다.
 *
 * 칸 값은 넷 중 하나다.
 *   null                      빈칸
 *   "문자"                    문자
 *   123                       숫자
 *   {kind:'date', iso, shown} 날짜 (iso = YYYY-MM-DD, shown = 엑셀 화면에 보이던 글자)
 */
(function (global) {
  'use strict';

  var XLSX = global.XLSX;

  // 엑셀을 읽을 때 쓰는 설정. cellNF 가 있어야 날짜 서식을 알아볼 수 있다.
  var READ_OPTS = { type: 'array', cellNF: true, cellDates: false, cellText: true };

  function CleanError(message) {
    this.name = 'CleanError';
    this.message = message;
  }
  CleanError.prototype = Object.create(Error.prototype);

  // ---------------------------------------------------------------- 값 다루기

  function isDate(v) { return v !== null && typeof v === 'object' && v.kind === 'date'; }
  function isBlank(v) { return v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); }
  function pad(n, w) { n = String(n); while (n.length < (w || 2)) n = '0' + n; return n; }
  function isoOf(y, m, d) { return pad(y, 4) + '-' + pad(m) + '-' + pad(d); }

  function realDate(y, m, d) {
    if (!(y >= 1000 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return false;
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
  }

  function serialOf(iso) {
    var p = iso.split('-').map(Number);
    return (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  // 글자로 적힌 날짜를 읽는다.
  // 돌려주는 것: {iso} · {ambiguous:true} (월/일을 가릴 수 없음) · null (읽지 못함)
  function parseDateText(text) {
    var t = String(text).trim();
    var m;
    function ymd(y, mo, d) { return realDate(y, mo, d) ? { iso: isoOf(y, mo, d) } : null; }
    if ((m = t.match(/^(\d{4})(\d{2})(\d{2})$/))) return ymd(+m[1], +m[2], +m[3]);
    if ((m = t.match(/^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일$/))) return ymd(+m[1], +m[2], +m[3]);
    if ((m = t.match(/^(\d{4})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{1,2})\s*\.?$/))) return ymd(+m[1], +m[2], +m[3]);
    if ((m = t.match(/^(\d{1,2})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{4})$/))) {
      var a = +m[1], b = +m[2], y = +m[3];
      var asMD = realDate(y, a, b), asDM = realDate(y, b, a);
      if (asMD && asDM && a !== b) return { ambiguous: true };
      if (asMD) return { iso: isoOf(y, a, b) };
      if (asDM) return { iso: isoOf(y, b, a) };
      return null;
    }
    return null;
  }

  // 화면과 기록에 보여 줄 글자
  function show(v, fmt) {
    if (v === null || v === undefined) return '';
    if (isDate(v)) return v.shown;
    if (typeof v === 'number') return fmt ? XLSX.SSF.format(fmt, v) : String(v);
    return String(v);
  }

  // 바뀌었는지 볼 때 쓰는 값 (모양까지 본다)
  function exact(v) {
    if (v === null || v === undefined) return 'E';
    if (isDate(v)) return 'D' + v.iso + '|' + v.shown;
    if (typeof v === 'number') return 'N' + v;
    return 'S' + v;
  }

  // 같은 줄인지 볼 때 쓰는 값 (내용만 본다)
  function content(v) {
    if (v === null || v === undefined || v === '') return 'E';
    if (isDate(v)) return 'D' + v.iso;
    if (typeof v === 'number') return 'N' + v;
    return 'S' + v;
  }

  // 엑셀 칸 하나를 값으로 읽는다
  function readCell(cell, date1904) {
    if (!cell || cell.t === 'z') return null;
    if (cell.t === 's' || cell.t === 'str') return cell.v === '' ? null : String(cell.v);
    if (cell.t === 'n') {
      if (cell.z && XLSX.SSF.is_date(cell.z)) {
        var p = XLSX.SSF.parse_date_code(cell.v, { date1904: !!date1904 });
        if (p && realDate(p.y, p.m, p.d)) {
          return { kind: 'date', iso: isoOf(p.y, p.m, p.d), shown: cell.w != null ? String(cell.w) : XLSX.SSF.format(cell.z, cell.v) };
        }
      }
      return cell.v;
    }
    if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
    if (cell.t === 'e') return cell.w != null ? String(cell.w) : '#ERROR';
    return cell.w != null ? String(cell.w) : String(cell.v);
  }

  // ---------------------------------------------------------------- 시트 읽기

  function workbookFrom(data, extraOpts) {
    var opts = {};
    for (var k in READ_OPTS) opts[k] = READ_OPTS[k];
    for (var k2 in (extraOpts || {})) opts[k2] = extraOpts[k2];
    return XLSX.read(data, opts);
  }

  function is1904(wb) {
    return !!(wb && wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  }

  // 시트 하나를 머리글과 줄로 읽는다. 첫 줄을 머리글로 본다.
  function readSheet(ws, date1904) {
    var out = { header: [], rows: [], skippedEmpty: 0 };
    if (!ws || !ws['!ref']) return out;
    var range = XLSX.utils.decode_range(ws['!ref']);
    var r0 = range.s.r;
    for (var c = range.s.c; c <= range.e.c; c++) {
      var hv = readCell(ws[XLSX.utils.encode_cell({ r: r0, c: c })], date1904);
      var name = hv === null ? '' : show(hv).trim();
      out.header.push(name || '(' + XLSX.utils.encode_col(c) + '열)');
    }
    for (var r = r0 + 1; r <= range.e.r; r++) {
      var values = [];
      var allBlank = true;
      for (var c2 = range.s.c; c2 <= range.e.c; c2++) {
        var v = readCell(ws[XLSX.utils.encode_cell({ r: r, c: c2 })], date1904);
        if (!isBlank(v)) allBlank = false;
        values.push(v);
      }
      if (allBlank) { out.skippedEmpty++; continue; }
      out.rows.push({ rowNum: r + 1, values: values });
    }
    return out;
  }

  function sheetSizes(wb) {
    return wb.SheetNames.map(function (name) {
      var s = readSheet(wb.Sheets[name], is1904(wb));
      return { name: name, rows: s.rows.length, cols: s.header.length, skippedEmpty: s.skippedEmpty };
    });
  }

  function pickSheet(wb, rules) {
    if (rules && rules.sheet && wb.SheetNames.indexOf(rules.sheet) >= 0) return { name: rules.sheet, warning: null };
    if (rules && rules.sheet && wb.SheetNames.length !== 1) {
      throw new CleanError('처방에 적힌 시트 «' + rules.sheet + '»가 이 파일에 없습니다. 이 파일의 시트: ' + wb.SheetNames.join(', '));
    }
    var warning = rules && rules.sheet ? '처방에 적힌 시트 «' + rules.sheet + '»가 없어서 하나뿐인 시트 «' + wb.SheetNames[0] + '»를 썼습니다.' : null;
    return { name: wb.SheetNames[0], warning: warning };
  }

  // 처방에 맞춰 표를 읽는다
  function readTable(wb, rules) {
    var pick = pickSheet(wb, rules);
    var sheet = readSheet(wb.Sheets[pick.name], is1904(wb));
    var flagName = rules.flag_column && rules.flag_column.name;
    var index = {};
    var dupHeaders = [];
    sheet.header.forEach(function (h, i) {
      if (index.hasOwnProperty(h)) dupHeaders.push(h); else index[h] = i;
    });
    var missing = rules.columns.filter(function (c) { return !index.hasOwnProperty(c); });
    var extra = sheet.header.filter(function (h, i) {
      return rules.columns.indexOf(h) < 0 && h !== flagName && index[h] === i;
    });
    var rows = sheet.rows.map(function (row) {
      var cells = {};
      rules.columns.forEach(function (c) { cells[c] = index.hasOwnProperty(c) ? row.values[index[c]] : null; });
      var ex = {};
      extra.forEach(function (c) { ex[c] = row.values[index[c]]; });
      return { rowNum: row.rowNum, cells: cells, extra: ex };
    });
    return {
      sheetName: pick.name,
      header: sheet.header,
      rows: rows,
      missingColumns: missing,
      extraColumns: extra,
      duplicateHeaders: dupHeaders.filter(function (h) { return rules.columns.indexOf(h) >= 0; }),
      skippedEmpty: sheet.skippedEmpty,
      warnings: pick.warning ? [pick.warning] : []
    };
  }

  // ---------------------------------------------------------------- 처방 확인

  var KNOWN_TESTS = ['blank', 'ambiguous_date', 'after_today', '>', '<', '=', '>=', '<=', '!='];

  function ruleColumns(rule) {
    if (Array.isArray(rule.columns)) return rule.columns;
    if (rule.column) return [rule.column];
    return [];
  }

  function expandConditions(conds) {
    var out = [];
    (conds || []).forEach(function (c) {
      var cols = Array.isArray(c.columns) ? c.columns : [c.column];
      cols.forEach(function (col) { out.push({ column: col, test: c.test, value: c.value }); });
    });
    return out;
  }

  // 처방 파일을 정리기가 실행할 수 있는지 본다. 문제 목록을 돌려준다.
  function validateRules(rules) {
    var errors = [];
    if (!rules || typeof rules !== 'object') return ['처방 파일이 JSON 객체가 아닙니다.'];
    if (!Array.isArray(rules.columns) || !rules.columns.length) errors.push('columns(칸 목록)가 없습니다.');
    if (!Array.isArray(rules.rules) || !rules.rules.length) errors.push('rules(처방 목록)가 없습니다.');
    if (!rules.key) errors.push('key(줄을 가리키는 칸)가 없습니다.');
    if (!rules.flag_column || !rules.flag_column.name) errors.push('flag_column.name(표시 칸 이름)이 없습니다.');
    var cols = rules.columns || [];
    (rules.rules || []).forEach(function (rule, i) {
      var label = '처방 ' + (i + 1) + ' «' + (rule.id || '이름 없음') + '»';
      if (!rule.id) errors.push(label + ': id가 없습니다.');
      if (!rule.why) errors.push(label + ': why(이유)가 비어 있습니다.');
      if (!ACTIONS.hasOwnProperty(rule.action)) { errors.push(label + ': 정리기가 모르는 action «' + rule.action + '»입니다.'); return; }
      ruleColumns(rule).concat(rule.compare || []).forEach(function (c) {
        if (cols.indexOf(c) < 0) errors.push(label + ': columns에 없는 칸 «' + c + '»를 씁니다.');
      });
      if (rule.action === 'map' && (!rule.map || typeof rule.map !== 'object')) errors.push(label + ': map이 없습니다.');
      if (rule.action === 'date_format' && rule.format !== 'YYYY-MM-DD') errors.push(label + ': 정리기는 날짜 형식 YYYY-MM-DD만 압니다.');
      if (rule.action === 'strip_unit' && !rule.unit) errors.push(label + ': unit이 없습니다.');
      if (rule.action === 'dedupe') {
        if (!Array.isArray(rule.compare) || !rule.compare.length) errors.push(label + ': compare(비교할 칸)가 없습니다.');
        if (rule.keep_rule !== 'first') errors.push(label + ': keep_rule은 first만 압니다.');
      }
      if (rule.action === 'flag') {
        if (!rule.check) { errors.push(label + ': check(조건)가 없습니다.'); return; }
        if (rule.check.match !== 'each' && rule.check.match !== 'all') errors.push(label + ': check.match는 each 또는 all이어야 합니다.');
        var conds = expandConditions(rule.check.conditions);
        if (!conds.length) errors.push(label + ': check.conditions가 비어 있습니다.');
        conds.forEach(function (c) {
          if (cols.indexOf(c.column) < 0) errors.push(label + ': columns에 없는 칸 «' + c.column + '»를 씁니다.');
          if (KNOWN_TESTS.indexOf(c.test) < 0) errors.push(label + ': 정리기가 모르는 조건 «' + c.test + '»입니다.');
          if (['>', '<', '=', '>=', '<=', '!='].indexOf(c.test) >= 0 && typeof c.value !== 'number') errors.push(label + ': 조건 «' + c.test + '»에 숫자 value가 없습니다.');
        });
      }
    });
    return errors;
  }

  // ---------------------------------------------------------------- 처방 실행

  function runTest(cond, v, ctx) {
    switch (cond.test) {
      case 'blank': return isBlank(v);
      case 'ambiguous_date': return typeof v === 'string' && !!(parseDateText(v) || {}).ambiguous;
      case 'after_today': return isDate(v) && v.iso > ctx.today;
      case '>': return typeof v === 'number' && v > cond.value;
      case '<': return typeof v === 'number' && v < cond.value;
      case '=': return typeof v === 'number' && v === cond.value;
      case '>=': return typeof v === 'number' && v >= cond.value;
      case '<=': return typeof v === 'number' && v <= cond.value;
      case '!=': return typeof v === 'number' && v !== cond.value;
    }
    throw new CleanError('정리기가 모르는 조건 «' + cond.test + '»입니다.');
  }

  var ACTIONS = {
    remove_spaces: function (rule, ctx) {
      ruleColumns(rule).forEach(function (col) {
        ctx.recs.forEach(function (rec) {
          var v = rec.cells[col];
          if (typeof v === 'string' && /\s/.test(v)) ctx.set(rec, col, v.replace(/\s+/g, ''), rule);
        });
      });
    },

    map: function (rule, ctx) {
      ruleColumns(rule).forEach(function (col) {
        ctx.recs.forEach(function (rec) {
          var v = rec.cells[col];
          if (typeof v === 'string' && Object.prototype.hasOwnProperty.call(rule.map, v)) ctx.set(rec, col, rule.map[v], rule);
        });
      });
    },

    date_format: function (rule, ctx) {
      ruleColumns(rule).forEach(function (col) {
        ctx.formats[col] = 'yyyy-mm-dd';
        ctx.recs.forEach(function (rec) {
          var v = rec.cells[col];
          if (isBlank(v)) return;
          if (isDate(v)) { ctx.set(rec, col, { kind: 'date', iso: v.iso, shown: v.iso }, rule); return; }
          var p = parseDateText(typeof v === 'number' ? String(v) : v);
          if (p && p.iso) { ctx.set(rec, col, { kind: 'date', iso: p.iso, shown: p.iso }, rule); return; }
          if (p && p.ambiguous && rule.skip_ambiguous) return;
          ctx.warn(rec, col, v, '날짜로 읽지 못해 그대로 두었습니다', rule);
        });
      });
    },

    strip_unit: function (rule, ctx) {
      ruleColumns(rule).forEach(function (col) {
        ctx.recs.forEach(function (rec) {
          var v = rec.cells[col];
          if (typeof v !== 'string') return;
          var t = v.trim();
          if (t.length <= rule.unit.length || t.slice(-rule.unit.length) !== rule.unit) return;
          var s = t.slice(0, -rule.unit.length).trim();
          ctx.set(rec, col, /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : s, rule);
        });
      });
    },

    to_number: function (rule, ctx) {
      ruleColumns(rule).forEach(function (col) {
        if (rule.display) ctx.formats[col] = rule.display;
        ctx.recs.forEach(function (rec) {
          var v = rec.cells[col];
          if (typeof v !== 'string' || isBlank(v)) return;
          var s = v;
          (rule.strip || []).forEach(function (tok) { s = s.split(tok).join(''); });
          s = s.trim();
          if (/^-?\d+(\.\d+)?$/.test(s)) ctx.set(rec, col, Number(s), rule);
          else ctx.warn(rec, col, v, '숫자로 바꾸지 못해 그대로 두었습니다', rule);
        });
      });
    },

    flag: function (rule, ctx) {
      var conds = expandConditions(rule.check.conditions);
      ctx.recs.forEach(function (rec) {
        var hit = conds.filter(function (c) { return runTest(c, rec.cells[c.column], ctx); });
        var cols = [];
        if (rule.check.match === 'each') cols = hit.map(function (c) { return c.column; });
        else if (hit.length === conds.length) cols = rule.check.flag_columns || conds.map(function (c) { return c.column; });
        if (cols.length) rec.flags.push({ ruleId: rule.id, columns: cols, why: rule.why });
      });
    },

    dedupe: function (rule, ctx) {
      var seen = {};
      ctx.recs = ctx.recs.filter(function (rec) {
        var k = rule.compare.map(function (c) { return content(rec.cells[c]); }).join('\u0001');
        if (seen.hasOwnProperty(k)) {
          ctx.removed.push({ rowNum: rec.rowNum, key: rec.key, keptRowNum: seen[k].rowNum, ruleId: rule.id, why: rule.why });
          return false;
        }
        seen[k] = rec;
        return true;
      });
    }
  };

  function todayIso() {
    var d = new Date();
    return isoOf(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  // 표 하나를 처방대로 정리한다. 원래 표는 바꾸지 않는다.
  function clean(table, rules, opts) {
    var problems = validateRules(rules);
    if (problems.length) throw new CleanError('처방 파일을 실행할 수 없습니다.\n- ' + problems.join('\n- '));
    if (table.missingColumns.length) throw new CleanError('이 파일에 처방이 쓰는 칸이 없습니다: ' + table.missingColumns.join(', '));
    if (table.duplicateHeaders.length) throw new CleanError('같은 이름의 칸이 두 개 있습니다: ' + table.duplicateHeaders.join(', '));

    var ctx = {
      today: (opts && opts.today) || todayIso(),
      formats: {},
      changes: [],
      warnings: [],
      removed: [],
      recs: table.rows.map(function (r) {
        var cells = {};
        for (var c in r.cells) cells[c] = r.cells[c];
        return { rowNum: r.rowNum, key: show(r.cells[rules.key]), cells: cells, extra: r.extra, flags: [], changedCols: {} };
      })
    };
    ctx.set = function (rec, col, nv, rule) {
      var old = rec.cells[col];
      if (exact(old) === exact(nv)) return;
      ctx.changes.push({ rowNum: rec.rowNum, key: rec.key, column: col, before: old, after: nv, ruleId: rule.id, why: rule.why });
      rec.cells[col] = nv;
      rec.changedCols[col] = true;
    };
    ctx.warn = function (rec, col, v, note, rule) {
      ctx.warnings.push({ rowNum: rec.rowNum, key: rec.key, column: col, value: show(v), note: note, ruleId: rule.id });
    };

    var before = ctx.recs.length;
    var steps = rules.rules.filter(function (r) { return r.action !== 'dedupe'; })
      .concat(rules.rules.filter(function (r) { return r.action === 'dedupe'; }));
    steps.forEach(function (rule) { ACTIONS[rule.action](rule, ctx); });

    var fc = rules.flag_column;
    var order = rules.columns;
    var kept = {};
    ctx.recs.forEach(function (rec) {
      kept[rec.rowNum] = true;
      var cols = [];
      rec.flags.forEach(function (f) { f.columns.forEach(function (c) { if (cols.indexOf(c) < 0) cols.push(c); }); });
      cols.sort(function (a, b) { return order.indexOf(a) - order.indexOf(b); });
      rec.flagText = cols.length ? (fc.message || '{칸}').split('{칸}').join(cols.join(fc.join || ', ')) : '';
    });

    var changes = ctx.changes.filter(function (ch) { return kept[ch.rowNum]; });
    var changedRows = {};
    changes.forEach(function (ch) { changedRows[ch.rowNum] = true; });
    var changedCells = {};
    changes.forEach(function (ch) { changedCells[ch.rowNum + '|' + ch.column] = true; });

    return {
      sheetName: table.sheetName,
      columns: rules.columns.slice(),
      extraColumns: table.extraColumns.slice(),
      flagName: fc.name,
      rows: ctx.recs,
      before: before,
      after: ctx.recs.length,
      changes: changes,
      changedCellCount: Object.keys(changedCells).length,
      changedRowCount: Object.keys(changedRows).length,
      flagged: ctx.recs.filter(function (r) { return r.flagText; }).map(function (r) {
        return {
          rowNum: r.rowNum, key: r.key, text: r.flagText,
          ruleIds: r.flags.map(function (f) { return f.ruleId; }),
          whys: r.flags.map(function (f) { return f.why; })
        };
      }),
      removed: ctx.removed,
      warnings: table.warnings.map(function (w) { return { rowNum: '', key: '', column: '', value: '', note: w, ruleId: '' }; }).concat(ctx.warnings),
      formats: ctx.formats,
      today: ctx.today,
      skippedEmpty: table.skippedEmpty
    };
  }

  // ---------------------------------------------------------------- 내려받을 파일

  function toCell(v, fmt) {
    if (v === null || v === undefined || v === '') return null;
    if (isDate(v)) return { t: 'n', v: serialOf(v.iso), z: fmt || 'yyyy-mm-dd' };
    if (typeof v === 'number') return fmt ? { t: 'n', v: v, z: fmt } : { t: 'n', v: v };
    return { t: 's', v: String(v) };
  }

  function widthOf(s) {
    var w = 0;
    String(s).split('').forEach(function (ch) { w += ch.charCodeAt(0) > 0x2e80 ? 2 : 1; });
    return w;
  }

  function sheetFromRows(header, rows) {
    var ws = {};
    var widths = header.map(function (h) { return widthOf(h); });
    header.forEach(function (h, c) { ws[XLSX.utils.encode_cell({ r: 0, c: c })] = { t: 's', v: h }; });
    rows.forEach(function (cells, i) {
      cells.forEach(function (cell, c) {
        if (!cell) return;
        ws[XLSX.utils.encode_cell({ r: i + 1, c: c })] = cell;
        var text = cell.z ? XLSX.SSF.format(cell.z, cell.v) : String(cell.v);
        widths[c] = Math.max(widths[c], widthOf(text));
      });
    });
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rows.length, 1), c: header.length - 1 } });
    if (!rows.length) ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: header.length - 1 } });
    ws['!cols'] = widths.map(function (w) { return { wch: Math.min(Math.max(w + 2, 6), 60) }; });
    return ws;
  }

  function textCells(values) {
    return values.map(function (v) {
      if (v === null || v === undefined || v === '') return null;
      return typeof v === 'number' ? { t: 'n', v: v } : { t: 's', v: String(v) };
    });
  }

  function cleanedHeader(result) {
    return result.columns.concat(result.extraColumns, [result.flagName]);
  }

  // 정리본: 원래 칸 + (처방에 없던 칸) + 표시 칸
  function buildCleanWorkbook(result) {
    var header = cleanedHeader(result);
    var rows = result.rows.map(function (rec) {
      return result.columns.map(function (c) { return toCell(rec.cells[c], result.formats[c]); })
        .concat(result.extraColumns.map(function (c) { return toCell(rec.extra[c]); }))
        .concat([rec.flagText ? { t: 's', v: rec.flagText } : null]);
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheetFromRows(header, rows), result.sheetName);
    return wb;
  }

  // 바꾼 기록: 요약 · 고친 값 · 표시한 행 · 지운 행 · 그대로 둔 값 · 처방
  function buildLogWorkbook(result, rules, meta) {
    var wb = XLSX.utils.book_new();
    var summary = [
      ['원본 파일', meta.fileName],
      ['시트', result.sheetName],
      ['처방 파일', meta.rulesName],
      ['처방을 정한 사람', rules.owner || ''],
      ['처방 고친 날', rules.updated || ''],
      ['정리한 날', result.today],
      ['정리 전 행', result.before],
      ['정리 후 행', result.after],
      ['지운 행', result.removed.length],
      ['고친 칸', result.changedCellCount],
      ['고친 값이 있는 행', result.changedRowCount],
      ['표시한 행', result.flagged.length],
      ['그대로 둔 값', result.warnings.length],
      ['두 번 돌려도 같은 결과', meta.idempotent === true ? '같음' : meta.idempotent === false ? '다름' : '확인 안 함'],
      ['파일 처리 위치', '이 컴퓨터의 브라우저 안 (어디로도 보내지 않음)']
    ];
    XLSX.utils.book_append_sheet(wb, sheetFromRows(['항목', '값'], summary.map(textCells)), '요약');

    var fmts = result.formats;
    XLSX.utils.book_append_sheet(wb, sheetFromRows(
      ['원래 행', rules.key, '칸', '바꾸기 전', '바꾼 후', '처방', '이유'],
      result.changes.map(function (ch) {
        return textCells([ch.rowNum, ch.key, ch.column, show(ch.before), show(ch.after, fmts[ch.column]), ch.ruleId, ch.why]);
      })), '고친 값');

    XLSX.utils.book_append_sheet(wb, sheetFromRows(
      ['원래 행', rules.key, result.flagName, '처방', '이유'],
      result.flagged.map(function (f) {
        return textCells([f.rowNum, f.key, f.text, f.ruleIds.join(', '), f.whys.join(' / ')]);
      })), '표시한 행');

    XLSX.utils.book_append_sheet(wb, sheetFromRows(
      ['원래 행', rules.key, '남긴 행', '처방', '이유'],
      result.removed.map(function (r) {
        return textCells([r.rowNum, r.key, r.keptRowNum, r.ruleId, r.why]);
      })), '지운 행');

    XLSX.utils.book_append_sheet(wb, sheetFromRows(
      ['원래 행', rules.key, '칸', '값', '까닭', '처방'],
      result.warnings.map(function (w) {
        return textCells([w.rowNum, w.key, w.column, w.value, w.note, w.ruleId]);
      })), '그대로 둔 값');

    XLSX.utils.book_append_sheet(wb, sheetFromRows(
      ['처방', '하는 일', '이유', '정한 사람'],
      rules.rules.map(function (r) { return textCells([r.id, r.what || '', r.why || '', r.owner || rules.owner || '']); })), '처방');
    return wb;
  }

  // ---------------------------------------------------------------- 두 번 돌리기

  function sameRows(a, b) {
    var diffs = [];
    if (a.rows.length !== b.rows.length) diffs.push('행 수가 다름: ' + a.rows.length + ' / ' + b.rows.length);
    var n = Math.min(a.rows.length, b.rows.length);
    for (var i = 0; i < n && diffs.length < 20; i++) {
      var x = a.rows[i], y = b.rows[i];
      a.columns.forEach(function (c) {
        if (content(x.cells[c]) !== content(y.cells[c]) || show(x.cells[c], a.formats[c]) !== show(y.cells[c], b.formats[c])) {
          diffs.push(x.key + ' ' + c + ': ' + show(x.cells[c], a.formats[c]) + ' / ' + show(y.cells[c], b.formats[c]));
        }
      });
      if (x.flagText !== y.flagText) diffs.push(x.key + ' ' + a.flagName + ': ' + x.flagText + ' / ' + y.flagText);
    }
    return diffs;
  }

  // 정리본을 실제 엑셀 파일로 만들어 다시 읽고, 한 번 더 정리해 본다.
  function checkIdempotent(result, rules, opts) {
    var bytes = XLSX.write(buildCleanWorkbook(result), { type: 'array', bookType: 'xlsx' });
    var wb2 = workbookFrom(bytes);
    var second = clean(readTable(wb2, rules), rules, { today: result.today });
    var diffs = sameRows(result, second);
    if (second.changes.length) diffs.unshift('두 번째에 고친 칸 ' + second.changes.length + '개 (예: ' + second.changes.slice(0, 3).map(function (c) { return c.key + ' ' + c.column; }).join(', ') + ')');
    if (second.removed.length) diffs.unshift('두 번째에 지운 행 ' + second.removed.length + '개');
    return { same: diffs.length === 0, diffs: diffs, second: second };
  }

  // ---------------------------------------------------------------- 정답과 대조

  function compareValue(v) {
    if (v === null || v === undefined) return '';
    if (isDate(v)) return v.iso;
    if (typeof v === 'number') return 'N:' + v;
    var s = String(v);
    if (s.trim() === '') return '';
    if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s) || /^-?\d+(\.\d+)?$/.test(s)) return 'N:' + Number(s.replace(/,/g, ''));
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m && realDate(+m[1], +m[2], +m[3])) return s;
    return 'S:' + s;
  }

  function compareWithAnswer(result, answerWb, rules) {
    var name = answerWb.SheetNames.indexOf(result.sheetName) >= 0 ? result.sheetName : answerWb.SheetNames[0];
    var ans = readSheet(answerWb.Sheets[name], is1904(answerWb));
    var ourHeader = cleanedHeader(result);
    var common = ourHeader.filter(function (h) { return ans.header.indexOf(h) >= 0; });
    var onlyOurCols = ourHeader.filter(function (h) { return ans.header.indexOf(h) < 0; });
    var onlyAnswerCols = ans.header.filter(function (h) { return ourHeader.indexOf(h) < 0; });

    function ourValue(rec, col) {
      if (col === result.flagName) return rec.flagText || null;
      if (result.columns.indexOf(col) >= 0) return rec.cells[col];
      return rec.extra[col];
    }
    function showOur(rec, col) { return show(ourValue(rec, col), result.formats[col]); }

    var keyCol = rules.key;
    var useKey = ans.header.indexOf(keyCol) >= 0 && ourHeader.indexOf(keyCol) >= 0;
    var ansRows = ans.rows.map(function (r) {
      var o = {};
      ans.header.forEach(function (h, i) { if (!o.hasOwnProperty(h)) o[h] = r.values[i]; });
      return { rowNum: r.rowNum, cells: o };
    });

    var pairs = [], onlyAnswer = [], onlyOurs = [], dupAnswerKeys = [];
    if (useKey) {
      var ours = {};
      result.rows.forEach(function (rec) { if (!ours.hasOwnProperty(rec.key)) ours[rec.key] = rec; });
      var used = {};
      ansRows.forEach(function (a) {
        var k = show(a.cells[keyCol]);
        if (used.hasOwnProperty(k)) { dupAnswerKeys.push(k); return; }
        used[k] = true;
        if (ours.hasOwnProperty(k)) pairs.push({ ours: ours[k], answer: a, key: k });
        else onlyAnswer.push({ rowNum: a.rowNum, key: k });
      });
      result.rows.forEach(function (rec) { if (!used.hasOwnProperty(rec.key)) onlyOurs.push({ rowNum: rec.rowNum, key: rec.key }); });
    } else {
      var n = Math.min(ansRows.length, result.rows.length);
      for (var i = 0; i < n; i++) pairs.push({ ours: result.rows[i], answer: ansRows[i], key: (i + 2) + '번째 줄' });
      ansRows.slice(n).forEach(function (a) { onlyAnswer.push({ rowNum: a.rowNum, key: a.rowNum + '행' }); });
      result.rows.slice(n).forEach(function (r) { onlyOurs.push({ rowNum: r.rowNum, key: r.key }); });
    }

    var mismatches = [];
    var matchedRows = 0;
    pairs.forEach(function (p) {
      var ok = true;
      common.forEach(function (col) {
        if (compareValue(ourValue(p.ours, col)) !== compareValue(p.answer.cells[col])) {
          ok = false;
          mismatches.push({ key: p.key, column: col, ours: showOur(p.ours, col), answer: show(p.answer.cells[col]), answerRow: p.answer.rowNum });
        }
      });
      if (ok) matchedRows++;
    });

    return {
      answerSheet: name,
      answerRows: ansRows.length,
      ourRows: result.rows.length,
      keyUsed: useKey ? keyCol : null,
      compared: pairs.length,
      matchedRows: matchedRows,
      mismatches: mismatches,
      onlyAnswer: onlyAnswer,
      onlyOurs: onlyOurs,
      dupAnswerKeys: dupAnswerKeys,
      commonColumns: common,
      onlyOurCols: onlyOurCols,
      onlyAnswerCols: onlyAnswerCols
    };
  }

  // ---------------------------------------------------------------- 처방 파일 찾기

  function baseName(name) {
    return String(name || '').replace(/\.[^.]+$/, '').replace(/\s*\(\d+\)$/, '').trim();
  }

  // ruleSets: [{name: '출장비정산.json', rules: {...}}]
  function findRules(fileName, ruleSets) {
    var b = baseName(fileName);
    for (var i = 0; i < ruleSets.length; i++) {
      var rs = ruleSets[i];
      if (baseName(rs.name) === b || (rs.rules && baseName(rs.rules.file) === b)) return rs;
    }
    return null;
  }

  global.Cleaner = {
    CleanError: CleanError,
    workbookFrom: workbookFrom,
    sheetSizes: sheetSizes,
    readTable: readTable,
    validateRules: validateRules,
    clean: clean,
    checkIdempotent: checkIdempotent,
    buildCleanWorkbook: buildCleanWorkbook,
    buildLogWorkbook: buildLogWorkbook,
    compareWithAnswer: compareWithAnswer,
    findRules: findRules,
    baseName: baseName,
    show: show,
    todayIso: todayIso
  };
})(window);
