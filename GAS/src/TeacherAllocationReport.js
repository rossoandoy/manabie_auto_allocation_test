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
  allocValues.slice(1).forEach(row => {
    const tId = row[o01Teacher] != null ? String(row[o01Teacher]).trim() : '';
    const sId = row[o01Student] != null ? String(row[o01Student]).trim() : '';
    const cId = row[o01Subject] != null ? String(row[o01Subject]).trim() : '';
    if (!tId || !sId || !cId) return;
    const key = `${tId}\t${sId}\t${cId}`;
    allocMap.set(key, (allocMap.get(key) || 0) + 1);
  });

  // 講師ごとの集計用
  const teacherSummary = new Map();
  const detailRows = [];

  planMap.forEach((planned, key) => {
    const [tId, sId, cId] = key.split('\t');
    const allocated = allocMap.get(key) || 0;
    const rate = planned > 0 ? Math.round((allocated / planned) * 1000) / 10 : 0;
    const tName = teacherNameMap.get(tId) || tId;
    const sName = studentNameMap.get(sId) || sId;
    const cName = subjectNameMap.get(cId) || cId;
    detailRows.push([tName, sName, cName, planned, allocated, rate]);

    if (!teacherSummary.has(tId)) {
      teacherSummary.set(tId, { teacherName: tName, planned: 0, allocated: 0, perStudent: new Map() });
    }
    const sum = teacherSummary.get(tId);
    sum.planned += planned;
    sum.allocated += allocated;
    sum.perStudent.set(sId, (sum.perStudent.get(sId) || 0) + allocated);
  });

  // 講師サマリ＋バランス指標
  const summaryRows = [];
  teacherSummary.forEach((sum, tId) => {
    const counts = Array.from(sum.perStudent.values());
    const n = counts.length;
    const mean = n > 0 ? counts.reduce((a, b) => a + b, 0) / n : 0;
    const variance = n > 0 ? counts.reduce((acc, c) => acc + Math.pow(c - mean, 2), 0) / n : 0;
    const stdDev = Math.sqrt(variance);
    const maxMinDiff = n > 1 ? Math.max(...counts) - Math.min(...counts) : 0;
    const balanceCv = mean > 0 ? Math.round((stdDev / mean) * 1000) / 10 : 0;
    const totalRate = sum.planned > 0 ? Math.round((sum.allocated / sum.planned) * 1000) / 10 : 0;
    summaryRows.push([
      sum.teacherName,
      n,
      sum.planned,
      sum.allocated,
      totalRate,
      Math.round(stdDev * 100) / 100,
      maxMinDiff,
      balanceCv
    ]);
  });
  summaryRows.sort((a, b) => (a[0] || '').localeCompare(b[0] || ''));

  // 出力シート
  const outSheetName = 'O04_teacher_allocation_report';
  let sheetOut = ss.getSheetByName(outSheetName);
  if (!sheetOut) sheetOut = ss.insertSheet(outSheetName); else sheetOut.clear();

  let row = 1;
  sheetOut.getRange(row, 1).setValue('【講師別サマリ】計画vs配置・バランス').setFontSize(12).setFontWeight('bold');
  row += 2;
  const summaryHeaders = ['講師名', '担当生徒数', '総計画コマ', '総配置コマ', '充足率(%)', 'バランス_標準偏差', 'バランス_最大最小差', 'バランス_変動係数(%)'];
  sheetOut.getRange(row, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  sheetOut.getRange(row, 1, 1, summaryHeaders.length).setBackground('#434343').setFontColor('white').setFontWeight('bold');
  row++;
  if (summaryRows.length > 0) {
    sheetOut.getRange(row, 1, summaryRows.length, summaryHeaders.length).setValues(summaryRows);
    sheetOut.getRange(row, 1, summaryRows.length, summaryHeaders.length).setBorder(true, true, true, true, true, true);
    row += summaryRows.length;
  }
  row += 2;
  sheetOut.getRange(row, 1).setValue('【明細】講師・生徒・科目の予定コマ数と配置コマ数の対応（講師でソート）').setFontSize(12).setFontWeight('bold');
  row += 2;
  const detailHeaders = ['講師名', '生徒名', '科目名', '予定コマ数', '配置コマ数', '充足率(%)'];
  sheetOut.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  sheetOut.getRange(row, 1, 1, detailHeaders.length).setBackground('#434343').setFontColor('white').setFontWeight('bold');
  row++;
  detailRows.sort(function (a, b) {
    const cmpTeacher = (a[0] || '').localeCompare(b[0] || '');
    if (cmpTeacher !== 0) return cmpTeacher;
    const cmpStudent = (a[1] || '').localeCompare(b[1] || '');
    if (cmpStudent !== 0) return cmpStudent;
    return (a[2] || '').localeCompare(b[2] || '');
  });
  if (detailRows.length > 0) {
    sheetOut.getRange(row, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
    sheetOut.getRange(row, 1, detailRows.length, detailHeaders.length).setBorder(true, true, true, true, true, true);
  }

  sheetOut.autoResizeColumns(1, Math.max(summaryHeaders.length, detailHeaders.length));
  ss.setActiveSheet(sheetOut);
  Browser.msgBox('講師別計画vs配置・バランス集計を出力しました。シート: ' + outSheetName);
}

function findColumnIndex(headerRow, candidates) {
  for (let i = 0; i < headerRow.length; i++) {
    const h = (headerRow[i] || '').toLowerCase();
    for (let j = 0; j < candidates.length; j++) {
      if (h.indexOf((candidates[j] || '').toLowerCase()) !== -1) return i;
    }
  }
  return -1;
}
