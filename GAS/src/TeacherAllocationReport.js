/**
 * 講師別：計画vs配置・バランス集計を表示
 * I07_student_subject（予定：生徒・科目・desired_teacher_1〜3 と各 max_slot で計画）と
 * O01_output_allocated_lessons（配置結果）を講師軸で集計し、
 * 計画数・配置数・充足率・バランス指標を出力する。
 */
function showTeacherAllocationReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetI07 = ss.getSheetByName('I07_student_subject');
  const sheetO01 = ss.getSheetByName('O01_output_allocated_lessons');
  const sheetStudents = ss.getSheetByName('I03_student_list');
  const sheetTeachers = ss.getSheetByName('I04_teacher_list');
  const sheetSubjects = ss.getSheetByName('I01_subject');

  if (!sheetI07 || sheetI07.getLastRow() < 2) {
    Browser.msgBox('エラー: I07_student_subject が見つからないか、データがありません。');
    return;
  }
  if (!sheetO01 || sheetO01.getLastRow() < 2) {
    Browser.msgBox('エラー: O01_output_allocated_lessons が見つからないか、データがありません。先にコマ組を実行してください。');
    return;
  }

  // 名前マッピング（ID -> 名前）
  const studentNameMap = new Map();
  const teacherNameMap = new Map();
  const subjectNameMap = new Map();
  if (sheetStudents && sheetStudents.getLastRow() >= 2) {
    sheetStudents.getRange(2, 1, sheetStudents.getLastRow(), 2).getValues().forEach(r => {
      if (r[0] != null && r[0] !== '') studentNameMap.set(String(r[0]).trim(), r[1] != null ? String(r[1]).trim() : '');
    });
  }
  if (sheetTeachers && sheetTeachers.getLastRow() >= 2) {
    sheetTeachers.getRange(2, 1, sheetTeachers.getLastRow(), 2).getValues().forEach(r => {
      if (r[0] != null && r[0] !== '') teacherNameMap.set(String(r[0]).trim(), r[1] != null ? String(r[1]).trim() : '');
    });
  }
  if (sheetSubjects && sheetSubjects.getLastRow() >= 2) {
    sheetSubjects.getRange(2, 1, sheetSubjects.getLastRow(), 2).getValues().forEach(r => {
      if (r[0] != null && r[0] !== '') subjectNameMap.set(String(r[0]).trim(), r[1] != null ? String(r[1]).trim() : '');
    });
  }

  // I07: desired_teacher_1〜3 と max_slot_1〜3（または max_slot）で計画を取得
  const numColsI07 = sheetI07.getLastColumn();
  const headerI07 = sheetI07.getRange(1, 1, 1, numColsI07).getValues()[0].map(h => String(h || '').trim());
  const idxStudent = findColumnIndex(headerI07, ['student_id', '生徒', '生徒ID']);
  const idxSubject = findColumnIndex(headerI07, ['subject_id', '科目', '科目ID']);
  const idxDesired1 = findColumnIndex(headerI07, ['desired_teacher_1']);
  const idxDesired2 = findColumnIndex(headerI07, ['desired_teacher_2']);
  const idxDesired3 = findColumnIndex(headerI07, ['desired_teacher_3']);
  const idxMaxSlot1 = findColumnIndex(headerI07, ['max_slot_1']);
  const idxMaxSlot2 = findColumnIndex(headerI07, ['max_slot_2']);
  const idxMaxSlot3 = findColumnIndex(headerI07, ['max_slot_3']);
  const idxMaxSlotAny = findColumnIndex(headerI07, ['max_slot', '予定コマ数', '計画コマ', 'sessions']);
  if (idxStudent < 0 || idxSubject < 0) {
    Browser.msgBox('エラー: I07_student_subject に student_id, subject_id 列が見つかりません。');
    return;
  }
  if (idxDesired1 < 0 && idxDesired2 < 0 && idxDesired3 < 0) {
    Browser.msgBox('エラー: I07_student_subject に desired_teacher_1〜3 のいずれかが見つかりません。');
    return;
  }
  const slotCol1 = idxMaxSlot1 >= 0 ? idxMaxSlot1 : idxMaxSlotAny;
  const slotCol2 = idxMaxSlot2 >= 0 ? idxMaxSlot2 : idxMaxSlotAny;
  const slotCol3 = idxMaxSlot3 >= 0 ? idxMaxSlot3 : idxMaxSlotAny;
  if (slotCol1 < 0 && slotCol2 < 0 && slotCol3 < 0) {
    Browser.msgBox('エラー: I07_student_subject に max_slot_1〜3 または max_slot に相当する列が見つかりません。');
    return;
  }

  const dataI07 = sheetI07.getRange(2, 1, sheetI07.getLastRow(), numColsI07).getValues();

  // 計画データ: (teacherId, studentId, subjectId) -> 予定コマ数（同一キーは合算）
  const planMap = new Map();
  dataI07.forEach(row => {
    const sId = row[idxStudent] != null ? String(row[idxStudent]).trim() : '';
    const cId = row[idxSubject] != null ? String(row[idxSubject]).trim() : '';
    if (!sId || !cId) return;
    const candidates = [
      { teacherCol: idxDesired1, slotCol: slotCol1 },
      { teacherCol: idxDesired2, slotCol: slotCol2 },
      { teacherCol: idxDesired3, slotCol: slotCol3 }
    ];
    candidates.forEach(function (c) {
      if (c.teacherCol < 0 || c.slotCol < 0) return;
      const tId = row[c.teacherCol] != null ? String(row[c.teacherCol]).trim() : '';
      if (!tId) return;
      const planned = parseInt(row[c.slotCol], 10) || 0;
      if (planned <= 0) return;
      const key = tId + '\t' + sId + '\t' + cId;
      planMap.set(key, (planMap.get(key) || 0) + planned);
    });
  });

  // O01: 配置数カウント (teacher_id, student_id, subject_id) -> 件数
  const allocValues = sheetO01.getDataRange().getValues();
  const headerO01 = allocValues[0].map(h => String(h || '').trim());
  const o01Teacher = headerO01.indexOf('teacher_id') >= 0 ? headerO01.indexOf('teacher_id') : headerO01.indexOf('講師ID');
  const o01Student = headerO01.indexOf('student_id') >= 0 ? headerO01.indexOf('student_id') : headerO01.indexOf('生徒ID');
  const o01Subject = headerO01.indexOf('subject_id') >= 0 ? headerO01.indexOf('subject_id') : headerO01.indexOf('科目ID');
  if (o01Teacher < 0 || o01Student < 0 || o01Subject < 0) {
    Browser.msgBox('エラー: O01 に teacher_id, student_id, subject_id 列が見つかりません。');
    return;
  }

  const allocMap = new Map();
  const o01SlotIdx = headerO01.indexOf('slot_id') >= 0 ? headerO01.indexOf('slot_id') : 0;
  const teacherSlotAllocs = {};
  allocValues.slice(1).forEach(function (row) {
    var tId = row[o01Teacher] != null ? String(row[o01Teacher]).trim() : '';
    var sId = row[o01Student] != null ? String(row[o01Student]).trim() : '';
    var cId = row[o01Subject] != null ? String(row[o01Subject]).trim() : '';
    var slotId = row[o01SlotIdx] != null ? String(row[o01SlotIdx]).trim() : '';
    if (!tId || !sId || !cId) return;
    var key = tId + '\t' + sId + '\t' + cId;
    allocMap.set(key, (allocMap.get(key) || 0) + 1);
    var slotKey = tId + '\t' + slotId;
    if (!teacherSlotAllocs[slotKey]) teacherSlotAllocs[slotKey] = [];
    teacherSlotAllocs[slotKey].push({ studentId: sId, subjectId: cId });
  });

  // I52: 講師の勤務可能コマ数（行数＝可能コマ数）と slot セット
  const teacherAvailableCount = new Map();
  const teacherSlotsSet = new Map();
  var sheetI52 = ss.getSheetByName('I52_teacher_availability');
  if (sheetI52 && sheetI52.getLastRow() >= 2) {
    var h52 = sheetI52.getRange(1, 1, 1, 2).getValues()[0].map(function (x) { return String(x || '').trim(); });
    var colT = h52[0] === 'teacher_id' || h52[0].indexOf('teacher') !== -1 ? 0 : 1;
    var colS = h52[1] === 'slot_id' || h52[1].indexOf('slot') !== -1 ? 1 : 0;
    sheetI52.getRange(2, 1, sheetI52.getLastRow(), 2).getValues().forEach(function (r) {
      var tid = String(r[colT] || '').trim();
      var sid = String(r[colS] || '').trim();
      if (!tid || !sid) return;
      teacherAvailableCount.set(tid, (teacherAvailableCount.get(tid) || 0) + 1);
      if (!teacherSlotsSet.has(tid)) teacherSlotsSet.set(tid, new Set());
      teacherSlotsSet.get(tid).add(sid);
    });
  }

  // I51: 生徒の希望スケジュール（slot_id のセット）
  const studentSlotsSet = new Map();
  var sheetI51 = ss.getSheetByName('I51_student_availability');
  if (sheetI51 && sheetI51.getLastRow() >= 2) {
    var h51 = sheetI51.getRange(1, 1, 1, 2).getValues()[0].map(function (x) { return String(x || '').trim(); });
    var colSt = h51[0] === 'student_id' || h51[0].indexOf('student') !== -1 ? 0 : 1;
    var colSl = h51[1] === 'slot_id' || h51[1].indexOf('slot') !== -1 ? 1 : 0;
    sheetI51.getRange(2, 1, sheetI51.getLastRow(), 2).getValues().forEach(function (r) {
      var stid = String(r[colSt] || '').trim();
      var sl = String(r[colSl] || '').trim();
      if (!stid || !sl) return;
      if (!studentSlotsSet.has(stid)) studentSlotsSet.set(stid, new Set());
      studentSlotsSet.get(stid).add(sl);
    });
  }

  // I06_teachable_subjects: 講師の指導可能科目 (teacher_id -> Set<subject_id>)
  const teacherTeachableSubjects = new Map();
  var sheetI06 = ss.getSheetByName('I06_teachable_subjects');
  if (sheetI06 && sheetI06.getLastRow() >= 2) {
    var h06 = sheetI06.getRange(1, 1, 1, 2).getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
    var colT = (h06[0] || '').indexOf('teacher') !== -1 ? 0 : 1;
    var colS = (h06[1] || '').indexOf('subject') !== -1 ? 1 : 0;
    sheetI06.getRange(2, 1, sheetI06.getLastRow(), 2).getValues().forEach(function (r) {
      var tid = String(r[colT] || '').trim();
      var sid = String(r[colS] || '').trim();
      if (!tid || !sid) return;
      if (!teacherTeachableSubjects.has(tid)) teacherTeachableSubjects.set(tid, new Set());
      teacherTeachableSubjects.get(tid).add(sid);
    });
  }

  // I05_lesson_slot と I02_time_range で slot_id -> 日付時限ラベル
  const slotIdToLabel = new Map();
  var sheetI05 = ss.getSheetByName('I05_lesson_slot');
  var sheetI02 = ss.getSheetByName('I02_time_range');
  if (sheetI05 && sheetI05.getLastRow() >= 2 && sheetI02 && sheetI02.getLastRow() >= 2) {
    var trMap = new Map();
    sheetI02.getRange(2, 1, sheetI02.getLastRow(), 2).getValues().forEach(function (r) {
      var id = r[0];
      var desc = r[1] != null ? String(r[1]).trim() : '';
      trMap.set(id, desc);
      trMap.set(String(id).trim(), desc);
    });
    sheetI05.getRange(2, 1, sheetI05.getLastRow(), 3).getValues().forEach(function (r) {
      var slotId = r[0];
      var dateVal = r[1];
      var timeId = r[2];
      if (slotId == null || slotId === '') return;
      var dateStr = dateVal ? Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), 'MM/dd') : '';
      var timeDesc = trMap.get(timeId) || trMap.get(String(timeId).trim()) || (timeId != null ? String(timeId) : '');
      var label = (dateStr && timeDesc) ? dateStr + ' ' + timeDesc : (dateStr || timeDesc || 'slot_' + slotId);
      slotIdToLabel.set(String(slotId).trim(), label);
      if (typeof slotId === 'number') slotIdToLabel.set(slotId, label);
    });
  }

  // 講師ごとの集計用（計画内の配置＋計画外の配置を分離）
  const teacherSummary = new Map();
  const detailRows = [];

  planMap.forEach(function (planned, key) {
    var parts = key.split('\t');
    var tId = parts[0], sId = parts[1], cId = parts[2];
    var allocated = allocMap.get(key) || 0;
    var rate = planned > 0 ? Math.round((allocated / planned) * 1000) / 10 : 0;
    var tName = teacherNameMap.get(tId) || tId;
    var sName = studentNameMap.get(sId) || sId;
    var cName = subjectNameMap.get(cId) || cId;
    var reason = inferUnallocatedReason(tId, sId, cId, planned, allocated, teacherSlotsSet, studentSlotsSet, teacherSlotAllocs, planMap, teacherAvailableCount, studentNameMap, slotIdToLabel, teacherTeachableSubjects);
    detailRows.push([tName, sName, cName, planned, allocated, rate, reason]);

    if (!teacherSummary.has(tId)) {
      teacherSummary.set(tId, { teacherName: tName, planned: 0, allocatedInPlan: 0, allocatedOutPlan: 0, perStudent: new Map() });
    }
    var sum = teacherSummary.get(tId);
    sum.planned += planned;
    sum.allocatedInPlan += allocated;
    sum.perStudent.set(sId, (sum.perStudent.get(sId) || 0) + allocated);
  });

  // O01 にのみある（I07計画外）講師・生徒・科目の明細を追加し、講師サマリに計画外配置を加算
  allocMap.forEach(function (allocated, key) {
    if (planMap.has(key)) return;
    var parts = key.split('\t');
    var tId = parts[0], sId = parts[1], cId = parts[2];
    var tName = teacherNameMap.get(tId) || tId;
    var sName = studentNameMap.get(sId) || sId;
    var cName = subjectNameMap.get(cId) || cId;
    detailRows.push([tName, sName, cName, 0, allocated, '-', 'I07計画外（担当外への配置）']);
    if (!teacherSummary.has(tId)) {
      teacherSummary.set(tId, { teacherName: tName, planned: 0, allocatedInPlan: 0, allocatedOutPlan: 0, perStudent: new Map() });
    }
    teacherSummary.get(tId).allocatedOutPlan += allocated;
  });

  // 講師サマリ＋勤務可能コマ数＋計画内/計画外/総配置＋バランス指標
  const summaryRows = [];
  teacherSummary.forEach(function (sum, tId) {
    var counts = Array.from(sum.perStudent.values());
    var n = counts.length;
    var mean = n > 0 ? counts.reduce(function (a, b) { return a + b; }, 0) / n : 0;
    var variance = n > 0 ? counts.reduce(function (acc, c) { return acc + Math.pow(c - mean, 2); }, 0) / n : 0;
    var stdDev = Math.sqrt(variance);
    var maxMinDiff = n > 1 ? Math.max.apply(null, counts) - Math.min.apply(null, counts) : 0;
    var balanceCv = mean > 0 ? Math.round((stdDev / mean) * 1000) / 10 : 0;
    var totalAlloc = (sum.allocatedInPlan || 0) + (sum.allocatedOutPlan || 0);
    var totalRate = sum.planned > 0 ? Math.round((sum.allocatedInPlan / sum.planned) * 1000) / 10 : 0;
    var availableCount = teacherAvailableCount.get(tId) || 0;
    summaryRows.push([
      sum.teacherName,
      n,
      availableCount,
      sum.planned,
      sum.allocatedInPlan || 0,
      sum.allocatedOutPlan || 0,
      totalAlloc,
      totalRate,
      Math.round(stdDev * 100) / 100,
      maxMinDiff,
      balanceCv
    ]);
  });
  summaryRows.sort(function (a, b) { return (a[0] || '').localeCompare(b[0] || ''); });

  // 出力シート
  const outSheetName = 'O04_teacher_allocation_report';
  let sheetOut = ss.getSheetByName(outSheetName);
  if (!sheetOut) sheetOut = ss.insertSheet(outSheetName); else sheetOut.clear();

  let row = 1;
  sheetOut.getRange(row, 1).setValue('【講師別サマリ】計画vs配置・バランス').setFontSize(12).setFontWeight('bold');
  row += 2;
  const summaryHeaders = ['講師名', '担当生徒数', '勤務可能コマ数', '総計画コマ', '計画内配置', '計画外配置', '総配置コマ', '充足率(%)', 'バランス_標準偏差', 'バランス_最大最小差', 'バランス_変動係数(%)'];
  sheetOut.getRange(row, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  sheetOut.getRange(row, 1, 1, summaryHeaders.length).setBackground('#434343').setFontColor('white').setFontWeight('bold');
  row++;
  if (summaryRows.length > 0) {
    sheetOut.getRange(row, 1, summaryRows.length, summaryHeaders.length).setValues(summaryRows);
    sheetOut.getRange(row, 1, summaryRows.length, summaryHeaders.length).setBorder(true, true, true, true, true, true);
    var summaryStartRow = row;
    for (var i = 0; i < summaryRows.length; i++) {
      var r = summaryStartRow + i;
      var avail = summaryRows[i][2];
      var plan = summaryRows[i][3];
      var allocIn = summaryRows[i][4];
      if (avail < plan) sheetOut.getRange(r, 4).setBackground('#f4c7c3');
      if (allocIn < plan) sheetOut.getRange(r, 5).setBackground('#f4c7c3');
    }
    row += summaryRows.length;
  }
  row += 2;
  sheetOut.getRange(row, 1).setValue('【明細】講師・生徒・科目の予定コマ数と配置コマ数の対応（講師でソート）').setFontSize(12).setFontWeight('bold');
  row += 2;
  const detailHeaders = ['講師名', '生徒名', '科目名', '予定コマ数', '配置コマ数', '充足率(%)', '未配置理由'];
  sheetOut.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  sheetOut.getRange(row, 1, 1, detailHeaders.length).setBackground('#434343').setFontColor('white').setFontWeight('bold');
  row++;
  detailRows.sort(function (a, b) {
    const cmpTeacher = (a[0] || '').localeCompare(b[0] || '');
    if (cmpTeacher !== 0) return cmpTeacher;
    const planA = (a[3] !== 0 && a[3] !== undefined && a[3] !== null) ? 0 : 1;
    const planB = (b[3] !== 0 && b[3] !== undefined && b[3] !== null) ? 0 : 1;
    if (planA !== planB) return planA - planB;
    const cmpStudent = (a[1] || '').localeCompare(b[1] || '');
    if (cmpStudent !== 0) return cmpStudent;
    return (a[2] || '').localeCompare(b[2] || '');
  });
  if (detailRows.length > 0) {
    sheetOut.getRange(row, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
    sheetOut.getRange(row, 1, detailRows.length, detailHeaders.length).setBorder(true, true, true, true, true, true);
    var detailStartRow = row;
    for (var d = 0; d < detailRows.length; d++) {
      var rate = detailRows[d][5];
      var detailRowNum = detailStartRow + d;
      var bg = null;
      if (typeof rate === 'number') {
        if (rate >= 100) {
          // 100%充足は色をつけない
        } else if (rate <= 25) {
          bg = '#f4c7c3';
        } else if (rate <= 50) {
          bg = '#fce8b2';
        } else if (rate <= 75) {
          bg = '#fff2cc';
        } else {
          bg = '#d9ead3';
        }
      }
      // 計画外（rate === '-'）は色なし
      if (bg) sheetOut.getRange(detailRowNum, 1, detailRowNum, detailHeaders.length).setBackground(bg);
    }
  }

  sheetOut.autoResizeColumns(1, Math.max(summaryHeaders.length, detailHeaders.length));
  ss.setActiveSheet(sheetOut);
  Browser.msgBox('講師別計画vs配置・バランス集計を出力しました。シート: ' + outSheetName);
}

/**
 * 未配置理由を推定（希望科目と指導可能科目の不一致、講師・生徒の希望日、競合、空きコマ不足など）
 * 競合時は 日付時限 と競合生徒名を「日付時限: 02/15 10:00, 02/16 11:00 | 競合生徒: A, B」形式で付加
 */
function inferUnallocatedReason(tId, sId, cId, planned, allocated, teacherSlotsSet, studentSlotsSet, teacherSlotAllocs, planMap, teacherAvailableCount, studentNameMap, slotIdToLabel, teacherTeachableSubjects) {
  if (allocated >= planned) return '充足';
  var teachable = teacherTeachableSubjects && teacherTeachableSubjects.get(tId);
  if (teachable && teachable.size > 0) {
    var cIdStr = String(cId || '').trim();
    if (cIdStr && !teachable.has(cIdStr)) return '希望科目と指導可能科目の不一致';
  }
  var tSlots = teacherSlotsSet.get(tId);
  var sSlots = studentSlotsSet.get(sId);
  if (!tSlots || tSlots.size === 0) return '講師の勤務可能データなし';
  if (!sSlots || sSlots.size === 0) return '生徒の希望日データなし';
  var common = [];
  tSlots.forEach(function (slot) {
    var slotStr = String(slot).trim();
    if (sSlots.has(slotStr) || sSlots.has(slot)) common.push(slotStr);
  });
  if (common.length === 0) return '講師の勤務日と生徒の希望日が不一致';
  var conflictSlots = [];
  var conflictStudentIds = {};
  for (var i = 0; i < common.length; i++) {
    var slotKey = tId + '\t' + common[i];
    var allocs = teacherSlotAllocs[slotKey];
    if (allocs) {
      for (var j = 0; j < allocs.length; j++) {
        if (allocs[j].studentId !== sId) {
          conflictSlots.push(common[i]);
          conflictStudentIds[allocs[j].studentId] = true;
          break;
        }
      }
    }
  }
  if (conflictSlots.length > 0) {
    var labels = [];
    for (var k = 0; k < conflictSlots.length && k < 20; k++) {
      var sid = conflictSlots[k];
      var lab = slotIdToLabel && (slotIdToLabel.get(sid) || slotIdToLabel.get(Number(sid)));
      labels.push(lab || ('slot_' + sid));
    }
    var dateTimeStr = labels.join(', ');
    if (conflictSlots.length > 20) dateTimeStr += '…';
    var names = [];
    for (var sid in conflictStudentIds) {
      names.push(studentNameMap && studentNameMap.get(sid) ? studentNameMap.get(sid) : sid);
    }
    names.sort();
    var nameStr = names.slice(0, 10).join(', ');
    if (names.length > 10) nameStr += ' 他' + (names.length - 10) + '名';
    return '同じ時限で他生徒と競合（日付時限: ' + dateTimeStr + ' | 競合生徒: ' + nameStr + '）';
  }
  var teacherTotalPlanned = 0;
  planMap.forEach(function (val, key) {
    if (key.indexOf(tId + '\t') === 0) teacherTotalPlanned += val;
  });
  var available = teacherAvailableCount.get(tId) || 0;
  if (available < teacherTotalPlanned) return '講師の空きコマ不足';
  return 'その他（制約に引っかかり）';
}

function findColumnIndex(headerRow, candidates) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = (headerRow[i] || '').toLowerCase();
    for (var j = 0; j < candidates.length; j++) {
      if (h.indexOf((candidates[j] || '').toLowerCase()) !== -1) return i;
    }
  }
  return -1;
}
