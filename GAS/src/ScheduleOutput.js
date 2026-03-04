function showScheduleExportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ScheduleDialog')
    .setWidth(420).setHeight(480).setTitle('印刷用スケジュールの作成');
  SpreadsheetApp.getUi().showModalDialog(html, '印刷用スケジュールの作成');
}

/**
 * マスタからリストを取得
 */
function getDropdownLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let students = []; let teachers = [];
  
  const studentSheet = ss.getSheetByName("I03_student_list");
  if (studentSheet) {
    const sData = studentSheet.getDataRange().getValues();
    for (let i = 1; i < sData.length; i++) {
      if (sData[i][1]) students.push(String(sData[i][1]).trim());
    }
  }

  const teacherSheet = ss.getSheetByName("I04_teacher_list");
  if (teacherSheet) {
    const tData = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      if (tData[i][1]) teachers.push(String(tData[i][1]).trim());
    }
  }

  return { students: Array.from(new Set(students)).sort(), teachers: Array.from(new Set(teachers)).sort() };
}

/**
 * IDベースのリレーションで組まれた堅牢な出力関数
 */
function generateScheduleMaster(params) {
  const { targetType, targetName, format } = params;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
  
  // 1. I02_time_range: time_range_id をキーにマスター化
  let periods = [];
  let trMap = {};
  const trSheet = ss.getSheetByName("I02_time_range");
  if (!trSheet) return { error: "I02_time_rangeシートが存在しません。" };
  const trData = trSheet.getDataRange().getValues();
  for (let i = 1; i < trData.length; i++) {
    let trId = String(trData[i][0]);
    let trName = String(trData[i][1]).trim();
    if (trId && trName) {
      trMap[trId] = { name: trName, index: periods.length };
      periods.push({ name: trName });
    }
  }
  const numPeriods = periods.length;
  if (numPeriods === 0) return { error: "時限マスターが空です。" };

  // 2. I05_lesson_slot: slot_id をキーに日付と時限情報をマスター化
  const slotSheet = ss.getSheetByName("I05_lesson_slot");
  if (!slotSheet) return { error: "I05_lesson_slotシートが存在しません。" };
  const slotData = slotSheet.getDataRange().getValues();
  
  let slotMap = {};
  let dateSet = new Set();
  let dateObjects = {};

  for (let i = 1; i < slotData.length; i++) {
    let slotId = String(slotData[i][0]);
    let rawDate = slotData[i][1];
    let trId = String(slotData[i][2]);

    if (slotId && rawDate && trMap[trId]) {
      let dObj = new Date(rawDate);
      if (!isNaN(dObj.getTime())) {
        let dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
        dateSet.add(dStr);
        if(!dateObjects[dStr]) dateObjects[dStr] = dObj; // 日付オブジェクトを保存
        
        // slot_idから「いつの、何限か」が即座に分かる辞書を作成
        slotMap[slotId] = {
          dateStr: dStr,
          periodIdx: trMap[trId].index
        };
      }
    }
  }
  
  let dateStrs = Array.from(dateSet).sort();
  if (dateStrs.length === 0) return { error: "I05_lesson_slotに有効な日付データが存在しません。" };
  
  let targetDates = dateStrs.map(dStr => dateObjects[dStr]);
  let titleStr = `講習会期間 (${Utilities.formatDate(targetDates[0], tz, "yyyy/M/d")} 〜 ${Utilities.formatDate(targetDates[targetDates.length-1], tz, "M/d")})`;

  // 3. O01_output_allocated_lessons: IDを頼りに確実にデータを抽出
  const allocatedSheet = ss.getSheetByName("O01_output_allocated_lessons");
  if (!allocatedSheet) return { error: "O01_output_allocated_lessonsシートが存在しません。" };
  const alData = allocatedSheet.getDataRange().getValues();
  
  let lessonMap = {}; // key: "YYYY-MM-DD_Index"
  const target = String(targetName).trim();

  for (let i = 1; i < alData.length; i++) {
    let row = alData[i];
    let slotId = String(row[0]);
    let student = String(row[5] || "").trim();
    let teacher = String(row[6] || "").trim();
    let subject = String(row[7] || "").trim();
    
    let slotInfo = slotMap[slotId];
    if (!slotInfo) continue; // 未知のスロットIDはスキップ
    
    let key = `${slotInfo.dateStr}_${slotInfo.periodIdx}`;
    
    if (targetType === 'student' && student === target) {
      if (!lessonMap[key]) lessonMap[key] = [];
      let teacherShort = teacher.split('(')[0];
      lessonMap[key].push({ subject: subject, partner: teacherShort });
    } 
    else if (targetType === 'teacher' && teacher === target) {
      if (!lessonMap[key]) lessonMap[key] = [];
      let studentShort = student.split('(')[0];
      lessonMap[key].push({ subject: subject, partner: studentShort });
    }
  }

  // 4. 出力用配列の構築（日付と時限の完全なマトリックス）
  let scheduleMap = []; 
  for (let i = 0; i < targetDates.length; i++) {
    let dateObj = targetDates[i];
    let m = dateObj.getMonth() + 1;
    let d = dateObj.getDate();
    let dow = dateObj.getDay();
    let dStr = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd");
    
    let dailyLessons = [];
    for (let pIdx = 0; pIdx < numPeriods; pIdx++) {
      let key = `${dStr}_${pIdx}`;
      if (lessonMap[key]) {
        let subjects = lessonMap[key].map(l => l.subject);
        let partners = lessonMap[key].map(l => l.partner);
        dailyLessons.push({ period: pIdx, col4: subjects.join(" / "), col5: partners.join(" / ") });
      }
    }
    scheduleMap.push({ month: m, day: d, dow: dow, lessons: dailyLessons });
  }

  // 5. 出力シートの準備と書き込み
  const printSheetName = format === 'vertical' ? "【印刷】タテ型リスト" : format === 'horizontal' ? "【印刷】ヨコ型リスト" : "【印刷】カレンダー形式";
  let printSheet = ss.getSheetByName(printSheetName);
  if (!printSheet) printSheet = ss.insertSheet(printSheetName, 0); 
  else { printSheet.clear(); printSheet.getRange(1, 1, printSheet.getMaxRows(), printSheet.getMaxColumns()).breakApart(); }

  const targetCleanName = targetName.split('(')[0]; 
  const honorific = targetType === 'student' ? "様" : "先生";

  // ==========================================
  // ★ タテ型リスト（縦に日付、横に時限）
  // ==========================================
  if (format === 'vertical') {
    printSheet.getRange("A1").setValue(`${targetCleanName} ${honorific}　${titleStr} 授業予定`).setFontSize(14).setFontWeight("bold");
    
    let headers = ["日付", "曜日"];
    for (let p = 0; p < numPeriods; p++) headers.push(periods[p].name);
    
    printSheet.getRange(3, 1, 1, headers.length).setValues([headers]).setBackground("#434343").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    printSheet.setColumnWidth(1, 50); printSheet.setColumnWidth(2, 50);
    let colWidth = numPeriods > 5 ? 100 : 130; 
    for (let p = 0; p < numPeriods; p++) printSheet.setColumnWidth(3 + p, colWidth);

    let outputData = [];
    scheduleMap.forEach(dObj => {
      let dowStr = dayNames[dObj.dow];
      let rowData = [`${dObj.month}/${dObj.day}`, dowStr];
      for (let p = 0; p < numPeriods; p++) {
        let lesson = dObj.lessons.find(l => l.period === p);
        if (lesson) rowData.push(`${lesson.col4}\n(${lesson.col5})`); else rowData.push("ー");
      }
      outputData.push(rowData);
    });

    if (outputData.length > 0) {
      let outRange = printSheet.getRange(4, 1, outputData.length, headers.length);
      outRange.setValues(outputData).setBorder(true, true, true, true, true, true).setWrap(true).setVerticalAlignment("middle");
      printSheet.getRange(4, 1, outputData.length, headers.length).setHorizontalAlignment("center");
      
      let currentRow = 4;
      scheduleMap.forEach((dObj, idx) => {
        let r = currentRow + idx;
        let rowRange = printSheet.getRange(r, 1, 1, headers.length);
        if (dObj.dow === 6) rowRange.setBackground("#eaf1fb");
        else if (dObj.dow === 0) rowRange.setBackground("#fce8e6");
        
        for (let p = 0; p < numPeriods; p++) {
          if (outputData[idx][2 + p] !== "ー") printSheet.getRange(r, 3 + p).setBackground("#fffbf0").setFontWeight("bold");
          else printSheet.getRange(r, 3 + p).setFontColor("#cccccc");
        }
        printSheet.setRowHeight(r, 45); 
      });
    }
  } 

  // ==========================================
  // ★ ヨコ型リスト（横に日付、縦に時限・14日分割）
  // ==========================================
  else if (format === 'horizontal') {
    printSheet.getRange("A1").setValue(`${targetCleanName} ${honorific}　${titleStr} 授業予定 (横印刷用)`).setFontSize(14).setFontWeight("bold");
    
    let chunks = [];
    let chunkSize = 14; // A4横に収まるように14日ごとに分割
    for (let i = 0; i < scheduleMap.length; i += chunkSize) chunks.push(scheduleMap.slice(i, i + chunkSize));

    let currentRow = 3;
    printSheet.setColumnWidth(1, 80); 
    for (let c = 2; c <= chunkSize + 1; c++) printSheet.setColumnWidth(c, 85);

    chunks.forEach((chunk) => {
      let dateRow = ["日付"]; let dowRow = ["曜日"];
      chunk.forEach(dObj => { dateRow.push(`${dObj.month}/${dObj.day}`); dowRow.push(dayNames[dObj.dow]); });

      printSheet.getRange(currentRow, 1, 1, dateRow.length).setValues([dateRow]).setBackground("#434343").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      currentRow++;
      
      printSheet.getRange(currentRow, 1, 1, dowRow.length).setValues([dowRow]).setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      chunk.forEach((dObj, i) => {
        if (dObj.dow === 6) printSheet.getRange(currentRow, i + 2).setBackground("#eaf1fb");
        if (dObj.dow === 0) printSheet.getRange(currentRow, i + 2).setBackground("#fce8e6");
      });
      currentRow++;

      let pRows = [];
      for (let p = 0; p < numPeriods; p++) {
        let pData = [periods[p].name];
        chunk.forEach(dObj => {
          let lesson = dObj.lessons.find(l => l.period === p);
          if (lesson) pData.push(`${lesson.col4}\n(${lesson.col5})`); else pData.push("ー");
        });
        pRows.push(pData);
      }

      let pRange = printSheet.getRange(currentRow, 1, numPeriods, pRows[0].length);
      pRange.setValues(pRows).setBorder(true, true, true, true, true, true).setWrap(true).setVerticalAlignment("middle").setHorizontalAlignment("center");

      for (let p = 0; p < numPeriods; p++) {
        printSheet.setRowHeight(currentRow + p, 45); 
        for (let i = 0; i < chunk.length; i++) {
          let targetCell = printSheet.getRange(currentRow + p, i + 2);
          if (pRows[p][i + 1] !== "ー") targetCell.setBackground("#fffbf0").setFontWeight("bold");
          else targetCell.setFontColor("#cccccc");
        }
      }
      currentRow += numPeriods + 2; 
    });
  }
  
  // ==========================================
  // ★ カレンダー形式（月またぎ完全対応）
  // ==========================================
  else if (format === 'calendar') {
    printSheet.getRange("A1").setValue(`${targetCleanName} ${honorific}　${titleStr} 授業カレンダー`).setFontSize(14).setFontWeight("bold");
    printSheet.getRange("A3:G3").setValues([["日", "月", "火", "水", "木", "金", "土"]]).setBackground("#434343").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    
    for (let c = 1; c <= 7; c++) printSheet.setColumnWidth(c, 110);
    
    let startCal = new Date(targetDates[0]);
    startCal.setDate(startCal.getDate() - startCal.getDay()); 
    let endCal = new Date(targetDates[targetDates.length - 1]);
    
    let gridData = [];
    let currentRow = 0; let currentCol = 0;
    let currentDate = new Date(startCal);
    
    while (currentDate <= endCal || currentDate.getDay() !== 0) {
      if (currentCol === 0) gridData.push(new Array(7).fill(""));
      
      let m = currentDate.getMonth() + 1;
      let d = currentDate.getDate();
      let dStr = Utilities.formatDate(currentDate, tz, "yyyy-MM-dd");
      
      let cellText = `${m}/${d}`; 
      let isTargetDay = targetDates.some(td => Utilities.formatDate(td, tz, "yyyy-MM-dd") === dStr);
      
      if (isTargetDay) {
        for (let pIdx = 0; pIdx < numPeriods; pIdx++) {
          let key = `${dStr}_${pIdx}`;
          if (lessonMap[key]) {
            let subjects = lessonMap[key].map(l => l.subject);
            let partners = lessonMap[key].map(l => l.partner);
            cellText += `\n[${periods[pIdx].name}] ${subjects.join(" / ")} (${partners.join(" / ")})`;
          }
        }
      }
      
      gridData[currentRow][currentCol] = cellText.trim();
      currentCol++;
      if (currentCol > 6) { currentCol = 0; currentRow++; }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    let outRange = printSheet.getRange(4, 1, gridData.length, 7);
    outRange.setValues(gridData).setBorder(true, true, true, true, true, true).setVerticalAlignment("top").setWrap(true);
    for (let r = 4; r < 4 + gridData.length; r++) printSheet.setRowHeight(r, 95);

    for (let r = 0; r < gridData.length; r++) {
      for (let c = 0; c < 7; c++) {
        let val = gridData[r][c];
        if (val) {
          if (c === 0) printSheet.getRange(r + 4, c + 1).setFontColor("#d93025");
          if (c === 6) printSheet.getRange(r + 4, c + 1).setFontColor("#1155cc");
          if (val.includes("[")) printSheet.getRange(r + 4, c + 1).setBackground("#fffbf0");
        }
      }
    }
  }

  ss.setActiveSheet(printSheet);
  return { success: true };
}