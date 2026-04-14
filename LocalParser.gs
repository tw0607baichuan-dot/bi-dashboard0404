/**
 * ============================================================
 *  LocalParser.gs — 本地班表矩阵解析器 v2.0
 * ============================================================
 *  适配真实结构：
 *  - R0 = 表头行（早班/中班/夜班 标题 + 各班人名）
 *  - R1 起 = 日期资料行
 *  - A/B 列 = 日期 / 星期（各区块前方也可能有日期列）
 *  - 用「早班/中班/夜班」标题切分区块（不依赖「值班人数」）
 * ============================================================
 */

var LocalParser = (function () {

  // ════════════════════════════════════
  //  STATUS_MAP — 格内文字 → 标准状态
  // ════════════════════════════════════
  var STATUS_MAP = {
    '班':     { code: 'on_duty',  category: 'active'    },
    '休':     { code: 'day_off',  category: 'off'       },
    '事假':   { code: 'leave',    category: 'leave'     },
    '特休':   { code: 'leave',    category: 'leave'     },
    '生理假': { code: 'leave',    category: 'leave'     },
    '迟':     { code: 'late',     category: 'exception' },
    '病假':   { code: 'leave',    category: 'leave'     }
  };

  // 班次标题关键字 → 内部 key 的对应
  var SHIFT_TITLES = [
    { keywords: ['早班'],           key: 'morning' },
    { keywords: ['中班'],           key: 'swing'   },
    { keywords: ['夜班', '晚班'],   key: 'night'   }
  ];

  // 表头中应跳过、不视为人名的标签
  var SKIP_LABELS = [
    '值班人数', '日期', '星期', '备注', '合计', '总计',
    '悦达本部', '早班', '中班', '夜班', '晚班',
    '连班人数', '出勤人数'
  ];

  // 底部汇总行终止词
  var TERMINATOR_KEYWORDS = [
    '悦达本部', '合计', '总计', '小计'
  ];

  // ════════════════════════════════════
  //  工具函数
  // ════════════════════════════════════

  function clean(val) {
    if (val === null || val === undefined) return '';
    // Google Sheets getValues() 对日期格回传 Date 物件 → 转 ISO 字串
    if (val instanceof Date) return val.toISOString();
    return String(val).trim();
  }

  function isDateOrNumber(str) {
    if (str === '') return false;
    // ISO 日期字串
    if (/^\d{4}-\d{2}-\d{2}T/.test(str)) return true;
    if (/^\d{1,4}$/.test(str)) return true;
    if (/^\d{1,4}[\/\-\.]\d{1,2}([\/\-\.]\d{1,4})?$/.test(str)) return true;
    return false;
  }

  function isSkipLabel(str) {
    if (str === '') return true;
    for (var i = 0; i < SKIP_LABELS.length; i++) {
      if (str === SKIP_LABELS[i]) return true;
    }
    if (isDateOrNumber(str)) return true;
    return false;
  }

  /**
   * 从一行中提取指定列范围内的人名及其列号
   */
  function extractNames(row, startCol, endCol) {
    var names = [];
    for (var c = startCol; c < endCol; c++) {
      var val = clean(row[c]);
      if (val !== '' && !isSkipLabel(val)) {
        names.push({ name: val, col: c });
      }
    }
    return names;
  }

  function parseDayNum(str) {
    if (str === '') return null;
    // ISO 日期字串：2026-04-01T07:00:00.000Z
    var isoMatch = str.match(/^\d{4}-\d{2}-(\d{2})T/);
    if (isoMatch) {
      var isoDay = parseInt(isoMatch[1], 10);
      if (isoDay >= 1 && isoDay <= 31) return isoDay;
    }
    // M/D 或 MM/DD
    var slashMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})$/);
    if (slashMatch) {
      var d = parseInt(slashMatch[2], 10);
      if (d >= 1 && d <= 31) return d;
    }
    // 纯数字 1~31
    var numMatch = str.match(/^(\d{1,2})$/);
    if (numMatch) {
      var n = parseInt(numMatch[1], 10);
      if (n >= 1 && n <= 31) return n;
    }
    return null;
  }

  function parseStatus(cellValue) {
    if (cellValue === '') return null;
    if (isDateOrNumber(cellValue)) return null;
    if (STATUS_MAP.hasOwnProperty(cellValue)) {
      var mapped = STATUS_MAP[cellValue];
      return { code: mapped.code, category: mapped.category };
    }
    return { code: 'unknown', category: 'unknown', raw: cellValue };
  }

  // ════════════════════════════════════
  //  核心定位函数（v2.1 — 班别标题 + 值班人数 精确框定）
  // ════════════════════════════════════

  /**
   * 在表头行中找到「早班」「中班」「夜班」所在的列号
   */
  function findShiftTitleCols(headerRow) {
    var result = { morning: -1, swing: -1, night: -1 };
    for (var c = 0; c < headerRow.length; c++) {
      var val = clean(headerRow[c]);
      if (val === '') continue;
      for (var si = 0; si < SHIFT_TITLES.length; si++) {
        var kws = SHIFT_TITLES[si].keywords;
        for (var ki = 0; ki < kws.length; ki++) {
          if (val.indexOf(kws[ki]) !== -1) {
            result[SHIFT_TITLES[si].key] = c;
          }
        }
      }
    }
    return result;
  }

  /**
   * 在表头行中找到所有「值班人数」的列号
   */
  function findDelimiterCols(headerRow) {
    var cols = [];
    for (var c = 0; c < headerRow.length; c++) {
      var val = clean(headerRow[c]);
      if (val === '值班人数' || val === '连班人数' || val === '出勤人数') {
        cols.push(c);
      }
    }
    return cols;
  }

  /**
   * 找到表头行：同时包含 ≥2 个班次标题 且 ≥1 个「值班人数」
   * 扫前 10 行
   */
  function findHeaderRow(data) {
    var scanLimit = Math.min(data.length, 10);
    for (var r = 0; r < scanLimit; r++) {
      var titles = findShiftTitleCols(data[r]);
      var found = 0;
      if (titles.morning !== -1) found++;
      if (titles.swing   !== -1) found++;
      if (titles.night   !== -1) found++;
      if (found < 2) continue;

      var delims = findDelimiterCols(data[r]);
      if (delims.length >= 1) return r;
    }
    return -1;
  }

  /**
   * 用班别标题 + 值班人数精确框定每个区块的人名列
   *
   * 实际布局（4月假表26 R1）：
   *   [早班] [空?] [name...] [值班人数] [中班] [空?] [name...] [值班人数] [夜班] [空?] [name...] [值班人数]
   *
   * 每个区块：标题列+1 ~ 该区块对应的「值班人数」列-1 之间提取人名
   */
  function buildShiftBlocks(headerRow, titleCols, delimCols) {
    // 按列号排序班次
    var order = [];
    if (titleCols.morning !== -1) order.push({ key: 'morning', col: titleCols.morning });
    if (titleCols.swing   !== -1) order.push({ key: 'swing',   col: titleCols.swing   });
    if (titleCols.night   !== -1) order.push({ key: 'night',   col: titleCols.night   });
    order.sort(function (a, b) { return a.col - b.col; });

    var result = { morning: [], swing: [], night: [] };
    var delimMap = {}; // key → delimCol

    for (var i = 0; i < order.length; i++) {
      var startCol = order[i].col + 1;
      // 找该区块后方最近的「值班人数」列作为 endCol
      var endCol = (i + 1 < order.length) ? order[i + 1].col : headerRow.length;
      for (var di = 0; di < delimCols.length; di++) {
        if (delimCols[di] > order[i].col && delimCols[di] < endCol) {
          endCol = delimCols[di]; // 精确截止于值班人数列
          delimMap[order[i].key] = delimCols[di];
          break;
        }
      }
      // 如果最后一个区块没有下一个标题，找其后方的 delim
      if (i === order.length - 1 && !delimMap[order[i].key]) {
        for (var di2 = 0; di2 < delimCols.length; di2++) {
          if (delimCols[di2] > order[i].col) {
            endCol = delimCols[di2];
            delimMap[order[i].key] = delimCols[di2];
            break;
          }
        }
      }
      result[order[i].key] = extractNames(headerRow, startCol, endCol);
    }

    result._delimMap = delimMap; // 保留供读取值班人数参考值
    return result;
  }

  /**
   * 定位资料起始行：从 headerRow+1 开始，找前 3 列含 Date 物件或可解析为小日期的行
   */
  function findDataStartRow(data, headerRowIdx) {
    for (var r = headerRowIdx + 1; r < data.length; r++) {
      for (var c = 0; c < Math.min(3, data[r].length); c++) {
        // 优先检测 Date 物件
        if (data[r][c] instanceof Date) return r;
        var d = parseDayNum(clean(data[r][c]));
        if (d !== null && d >= 1 && d <= 5) return r;
      }
    }
    return -1;
  }

  /**
   * 定位资料结束行
   */
  function findDataEndRow(data, startRow) {
    var emptyCount = 0;
    var lastValidRow = startRow;

    for (var r = startRow; r < data.length; r++) {
      var row = data[r];
      var firstVal = clean(row[0] || '');

      // 终止词检查（前 3 列）
      for (var c = 0; c < Math.min(3, row.length); c++) {
        var cv = clean(row[c] || '');
        for (var t = 0; t < TERMINATOR_KEYWORDS.length; t++) {
          if (cv.indexOf(TERMINATOR_KEYWORDS[t]) !== -1) {
            return lastValidRow;
          }
        }
      }

      if (firstVal === '') {
        emptyCount++;
        if (emptyCount >= 3) return lastValidRow;
      } else {
        emptyCount = 0;
        lastValidRow = r;
      }
    }
    return lastValidRow;
  }

  // ════════════════════════════════════
  //  主函数
  // ════════════════════════════════════

  function parse(sheet, options) {
    options = options || {};
    var month = options.month || null;
    var year  = options.year  || null;

    var data = sheet.getDataRange().getValues();
    var sheetName = sheet.getName();

    // 从 sheet 名称推断月份/年份
    if (!month) {
      var monthMatch = sheetName.match(/(\d{1,2})月/);
      if (monthMatch) month = parseInt(monthMatch[1], 10);
    }
    if (!year) {
      var yearMatch = sheetName.match(/(\d{4})/);
      if (yearMatch) year = parseInt(yearMatch[1], 10);
      else year = new Date().getFullYear();
    }
    var daysInMonth = month ? new Date(year, month, 0).getDate() : 31;

    var warnings = [];

    // ── 步骤 1：找表头行（含班次标题）──
    var headerRowIdx = findHeaderRow(data);
    if (headerRowIdx === -1) {
      return {
        type: 'local', error: 'HEADER_NOT_FOUND',
        message: '无法定位表头行：前 10 行中找不到包含 ≥2 个班次标题（早班/中班/夜班）的行',
        sheetName: sheetName, warnings: warnings
      };
    }
    var headerRow = data[headerRowIdx];
    var titleCols = findShiftTitleCols(headerRow);

    // ── 步骤 2：用班别标题 + 值班人数 精确切分区块 ──
    var delimCols = findDelimiterCols(headerRow);
    var shiftBlocks = buildShiftBlocks(headerRow, titleCols, delimCols);
    var delimMap = shiftBlocks._delimMap || {};

    // ── 步骤 3：定位资料起始行 ──
    var dataStartRow = findDataStartRow(data, headerRowIdx);
    if (dataStartRow === -1) {
      return {
        type: 'local', error: 'DATA_START_NOT_FOUND',
        message: '无法定位资料起始行：表头行之后找不到日期值',
        sheetName: sheetName, warnings: warnings
      };
    }

    // ── 步骤 4：定位资料结束行 ──
    var dataEndRow = findDataEndRow(data, dataStartRow);

    // ── 步骤 5：锁定日期列，然后逐行逐格解析 ──
    var daily = {};
    var refCounts = {};
    var shiftKeys = ['morning', 'swing', 'night'];
    var shiftNamesList = [shiftBlocks.morning, shiftBlocks.swing, shiftBlocks.night];

    // 5a. 先从前几行侦测「日期列」（只信任 Date 物件所在的列）
    var dateCol = -1;
    for (var probe = dataStartRow; probe < Math.min(dataStartRow + 5, data.length); probe++) {
      for (var pc = 0; pc < Math.min(3, data[probe].length); pc++) {
        if (data[probe][pc] instanceof Date) { dateCol = pc; break; }
      }
      if (dateCol !== -1) break;
    }
    // fallback：如果没侦测到 Date 物件，用 column 0
    if (dateCol === -1) dateCol = 0;

    var _parseLog = [];
    _parseLog.push('dateCol=' + dateCol);

    // 5b. 逐行解析，只从 dateCol 取日期
    for (var r = dataStartRow; r <= dataEndRow; r++) {
      var row = data[r];

      // 只从锁定的日期列取日期
      var rawDateVal = row[dateCol];
      var dayNum = null;

      if (rawDateVal instanceof Date) {
        // Date 物件：直接取本地日期（避免 UTC 偏移）
        dayNum = rawDateVal.getDate();
      } else {
        // 非 Date：用 parseDayNum 解析（仅限 dateCol）
        dayNum = parseDayNum(clean(rawDateVal));
      }

      if (dayNum === null) continue;
      if (dayNum < 1 || dayNum > daysInMonth) continue;

      var dayKey = String(dayNum);

      // 重复 dayKey 保护：不覆盖已有资料
      if (daily.hasOwnProperty(dayKey)) {
        _parseLog.push('⚠ SKIP duplicate dayKey=' + dayKey + ' at R' + r + ' (raw=' + JSON.stringify(rawDateVal).substring(0,40) + ')');
        continue;  // 跳过，不覆盖
      }

      _parseLog.push('R' + r + ' → day=' + dayKey);
      daily[dayKey] = {};
      refCounts[dayKey] = {};

      for (var si = 0; si < shiftKeys.length; si++) {
        var shiftKey = shiftKeys[si];
        var names = shiftNamesList[si];
        daily[dayKey][shiftKey] = {};

        for (var ni = 0; ni < names.length; ni++) {
          var staffName = names[ni].name;
          var staffCol  = names[ni].col;
          var cellVal   = clean(row[staffCol]);
          var status    = parseStatus(cellVal);

          daily[dayKey][shiftKey][staffName] = status;

          if (status !== null && status.code === 'unknown') {
            warnings.push({
              day: dayNum, shift: shiftKey, name: staffName,
              raw: status.raw, message: '未识别的状态值'
            });
          }
        }

        // 读取该班次的「值班人数」参考值
        var refCol = delimMap[shiftKey];
        if (refCol !== undefined && refCol < row.length) {
          var refVal = clean(row[refCol]);
          var refNum = parseInt(refVal, 10);
          refCounts[dayKey][shiftKey] = isNaN(refNum) ? null : refNum;
        }
      }
    }

    // ── 步骤 6：由 daily 重算 summary ──
    var summary = {};
    var dayKeys = Object.keys(daily);
    for (var di = 0; di < dayKeys.length; di++) {
      var dk = dayKeys[di];
      summary[dk] = {};
      for (var si2 = 0; si2 < shiftKeys.length; si2++) {
        var sk = shiftKeys[si2];
        var staffStatuses = daily[dk][sk] || {};
        var counts = { on_duty: 0, day_off: 0, leave: 0, exception: 0, 'null': 0, unknown: 0 };

        var staffNames = Object.keys(staffStatuses);
        for (var sni = 0; sni < staffNames.length; sni++) {
          var st = staffStatuses[staffNames[sni]];
          if (st === null) {
            counts['null']++;
          } else if (st.category === 'active') {
            counts.on_duty++;
          } else if (st.category === 'off') {
            counts.day_off++;
          } else if (st.category === 'leave') {
            counts.leave++;
          } else if (st.category === 'exception') {
            counts.exception++;
          } else {
            counts.unknown++;
          }
        }
        summary[dk][sk] = counts;
      }
    }

    // ── 组装输出 ──
    return {
      type: 'local',
      parsedAt: new Date().toISOString(),
      sheetName: sheetName,
      month: month,
      year: year,
      daysInMonth: daysInMonth,

      shifts: {
        morning: { names: shiftBlocks.morning.map(function(n) { return n.name; }), time: null },
        swing:   { names: shiftBlocks.swing.map(function(n) { return n.name; }),   time: null },
        night:   { names: shiftBlocks.night.map(function(n) { return n.name; }),   time: null }
      },

      daily: daily,
      summary: summary,
      refCounts: refCounts,
      warnings: warnings,
      _parseLog: _parseLog
    };
  }

  // ════════════════════════════════════
  //  测试 / 除错
  // ════════════════════════════════════

  /** 目标 sheet 名称 — 改这里即可切换 */
  var TARGET_SHEET = '4月假表26';

  /**
   * 依名称取得 sheet，找不到时 log 错误并回传 null
   */
  function getTargetSheet(sheetName) {
    var name = sheetName || TARGET_SHEET;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log('❌ 找不到 Sheet：「' + name + '」，请确认名称是否正确');
      return null;
    }
    return sheet;
  }

  function testParse(sheetName) {
    var targetSheet = getTargetSheet(sheetName);
    if (!targetSheet) return;

    Logger.log('解析 Sheet: ' + targetSheet.getName());
    var result = parse(targetSheet);
    Logger.log(JSON.stringify(result, null, 2));

    if (result.error) {
      Logger.log('❌ 解析失败: ' + result.error + ' — ' + result.message);
    } else {
      Logger.log('✅ 解析完成');
      Logger.log('   班次人数: 早=' + result.shifts.morning.names.length +
                 ' 中=' + result.shifts.swing.names.length +
                 ' 夜=' + result.shifts.night.names.length);
      Logger.log('   解析天数: ' + Object.keys(result.daily).length);
      Logger.log('   warnings: ' + result.warnings.length);
      if (result.warnings.length > 0) {
        Logger.log('   前 5 个 warning:');
        for (var w = 0; w < Math.min(5, result.warnings.length); w++) {
          var wn = result.warnings[w];
          Logger.log('     Day ' + wn.day + ' ' + wn.shift + ' ' + wn.name + ': "' + wn.raw + '"');
        }
      }
    }
    return result;
  }

  function debugDump(sheetName) {
    var targetSheet = getTargetSheet(sheetName);
    if (!targetSheet) return;

    var data = targetSheet.getDataRange().getValues();
    var rows = Math.min(data.length, 12);
    Logger.log('Sheet: ' + targetSheet.getName() + '  总行=' + data.length + '  总列=' + (data[0] ? data[0].length : 0));

    for (var r = 0; r < rows; r++) {
      var cells = [];
      for (var c = 0; c < data[r].length; c++) {
        var v = data[r][c];
        if (v === '' || v === null || v === undefined) continue;
        cells.push('C' + c + '=' + JSON.stringify(v));
      }
      Logger.log('R' + r + ': ' + cells.join(' | '));
    }
  }

  /**
   * 诊断：dump 指定日期那行的 raw cell 值 + 解析结果
   * 在 Apps Script 编辑器跑 runDebugDay() 即可
   */
  function debugDay(targetDay, sheetName) {
    targetDay = targetDay || 10;
    var targetSheet = getTargetSheet(sheetName);
    if (!targetSheet) return;

    var data = targetSheet.getDataRange().getValues();

    // 找表头
    var headerRowIdx = findHeaderRow(data);
    if (headerRowIdx === -1) { Logger.log('❌ 找不到表头行'); return; }
    var headerRow = data[headerRowIdx];
    var titleCols = findShiftTitleCols(headerRow);
    var delimCols = findDelimiterCols(headerRow);
    var shiftBlocks = buildShiftBlocks(headerRow, titleCols, delimCols);

    Logger.log('headerRow=' + headerRowIdx);
    Logger.log('morning names: ' + shiftBlocks.morning.map(function(n){return n.name+'@C'+n.col}).join(', '));

    // 找目标日期的行
    var dataStartRow = findDataStartRow(data, headerRowIdx);
    Logger.log('dataStartRow=' + dataStartRow);

    for (var r = dataStartRow; r < data.length; r++) {
      var row = data[r];
      var dayNum = null;
      for (var c = 0; c < Math.min(3, row.length); c++) {
        dayNum = parseDayNum(clean(row[c]));
        if (dayNum !== null) break;
      }
      if (dayNum !== targetDay) continue;

      // 找到目标行
      Logger.log('═══ Day ' + targetDay + ' found at row ' + r + ' ═══');

      // dump 该行所有非空格
      var cells = [];
      for (var c2 = 0; c2 < row.length; c2++) {
        var raw = row[c2];
        if (raw === '' || raw === null || raw === undefined) continue;
        cells.push('C' + c2 + '=' + JSON.stringify(raw) + ' (type=' + typeof raw + ')');
      }
      Logger.log('RAW: ' + cells.join(' | '));

      // 逐人解析
      var shiftKeys = ['morning', 'swing', 'night'];
      var shiftNamesList = [shiftBlocks.morning, shiftBlocks.swing, shiftBlocks.night];
      for (var si = 0; si < shiftKeys.length; si++) {
        var names = shiftNamesList[si];
        var results = [];
        for (var ni = 0; ni < names.length; ni++) {
          var col = names[ni].col;
          var rawVal = row[col];
          var cleaned = clean(rawVal);
          var status = parseStatus(cleaned);
          results.push(names[ni].name + '@C' + col + ': raw=' + JSON.stringify(rawVal) + '(' + typeof rawVal + ') clean="' + cleaned + '" status=' + JSON.stringify(status));
        }
        Logger.log(shiftKeys[si] + ':\n  ' + results.join('\n  '));
      }
      return; // 只看这一天
    }
    Logger.log('❌ Day ' + targetDay + ' not found in data rows');
  }

  return {
    parse: parse,
    testParse: testParse,
    debugDump: debugDump,
    debugDay: debugDay,
    TARGET_SHEET: TARGET_SHEET,
    STATUS_MAP: STATUS_MAP,
    SKIP_LABELS: SKIP_LABELS
  };

})();

// ═══ 顶层入口（Apps Script 编辑器可直接执行）═══
function runTestParse() { return LocalParser.testParse(); }
function runDebugDump()  { LocalParser.debugDump(); }
function runDebugDay()   { LocalParser.debugDay(10); }
// 如需指定其他 sheet，改上方 TARGET_SHEET 即可
