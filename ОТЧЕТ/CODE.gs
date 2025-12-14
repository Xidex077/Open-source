function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Создать отчет")
    .addItem("Сформировать еженедельную сводку", "createWeeklyReportFromMenu_")
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function createWeeklyReportFromMenu_() {
  try {
    runWeeklySummaryPro();
    SpreadsheetApp.getUi().alert("Отчет создан.");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Ошибка при создании отчета:\n" + err);
    throw err;
  }
}

function runWeeklySummaryPro() {
  removeAfterDash_();
  buildWeeklySummaryPro_();
}

function removeAfterDash_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("DATA");
  if (!sh) throw new Error('Лист "DATA" не найден');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const r = sh.getRange(2, 1, lastRow - 1, 1);
  const v = r.getValues();

  const out = v.map(row => {
    const s = row[0];
    if (s && typeof s === "string") {
      const i = s.indexOf("-");
      if (i !== -1) return [s.substring(0, i).trim()];
    }
    return row;
  });

  r.setValues(out);
}

function buildWeeklySummaryPro_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const today = new Date();
  const end = new Date(today);
  end.setDate(end.getDate() - 1);

  const start = new Date(end);
  start.setDate(start.getDate() - 6);

  const startStr = Utilities.formatDate(start, tz, "dd.MM");
  const endStr = Utilities.formatDate(end, tz, "dd.MM");
  const sheetName = `Сводка ${startStr}-${endStr}`;

  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const sh = ss.insertSheet(sheetName);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);

  const TOTAL_COLS = 12;
  sh.setColumnWidths(1, TOTAL_COLS, 110);

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 110);

  sh.setColumnWidth(6, 26);
  sh.setColumnWidth(7, 26);

  sh.setColumnWidth(8, 280);
  sh.setColumnWidth(9, 90);
  sh.setColumnWidth(10, 130);
  sh.setColumnWidth(11, 110);
  sh.setColumnWidth(12, 110);

  sh.setRowHeight(1, 64);
  const title = sh.getRange(1, 1, 1, TOTAL_COLS);
  title.merge();
  title
    .setValue("ЕЖЕНЕДЕЛЬНАЯ СВОДКА")
    .setFontFamily("Inter")
    .setFontSize(24)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#0F172A");
  applyOuterThick_(title);

  sh.setRowHeight(2, 30);
  const desc = sh.getRange(2, 1, 1, TOTAL_COLS);
  desc.merge();
  desc
    .setValue("ГРЯЗНЫЕ ПОТЕРИ ОТЧЕТА <Потери сборки и сортировки>")
    .setFontFamily("Inter")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#1F2937");
  applyOuterThick_(desc);

  sh.setRowHeight(3, 10);

  const totalRow = buildGeneralWriteoffsBlock_(ss, sh, 4);

  const spacerRow = totalRow + 1;
  sh.setRowHeight(spacerRow, 12);

  const sectionHeaderRow = spacerRow + 1;
  buildTwoSectionHeaders_(sh, sectionHeaderRow);

  const startTablesRow = sectionHeaderRow + 2;
  sh.setRowHeight(startTablesRow - 1, 10);

  const norm = makeNormalizer_();

  const endLeftBlocks = buildLossesByBlockTableSide_(ss, sh, startTablesRow, {
    startCol: 1,
    srcAllow: new Set([norm("Потери сортировки"), norm("Потери СЦ")])
  });

  const endRightBlocks = buildLossesByBlockTableSide_(ss, sh, startTablesRow, {
    startCol: 8,
    srcAllow: new Set([norm("Потери сборки")])
  });

  let leftRow = endLeftBlocks;
  let rightRow = endRightBlocks;

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildSortingMxBucketsTableLeft_(ss, sh, leftRow + 2);

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildTop10MxForPoterySCLeft_(ss, sh, leftRow + 2);

  const tokens = makeTokens_();

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildTop10MxForPoterySortirovkiFilteredLeft_(ss, sh, leftRow + 2, {
    title: "МХ ПОТОК",
    mxTransform: formatMxPotokToEtazh_,
    predicate: (cN, containsAuto) => (cN.indexOf(tokens.tokenPotok) !== -1) || (cN.indexOf(tokens.tokenBuf) !== -1)
  });

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildTop10MxForPoterySortirovkiFilteredLeft_(ss, sh, leftRow + 2, {
    title: "ПРЕДСОРТ",
    predicate: (cN, containsAuto) => (cN.indexOf(tokens.tokenPS) !== -1)
  });

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildTop10MxForPoterySortirovkiFilteredLeft_(ss, sh, leftRow + 2, {
    title: "СОРТ",
    predicate: (cN, containsAuto) => {
      const has = (cN.indexOf(tokens.tokenKS) !== -1) || (cN.indexOf(tokens.tokenSort) !== -1);
      return has && !containsAuto;
    }
  });

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildTop10MxForPoterySortirovkiFilteredLeft_(ss, sh, leftRow + 2, {
    title: "АВТОСОРТ",
    predicate: (cN, containsAuto) => containsAuto
  });

  sh.setRowHeight(leftRow + 1, 10);
  leftRow = buildTop10FioSkladchikaForPoterySortirovkiLeft_(ss, sh, leftRow + 2);

  sh.setRowHeight(rightRow + 1, 10);
  rightRow = buildTopEtazhPoterySborkiRight_(ss, sh, rightRow + 2);

  sh.setRowHeight(rightRow + 1, 10);
  rightRow = buildTopRyadPoterySborkiRight_(ss, sh, rightRow + 2);

  sh.setRowHeight(rightRow + 1, 10);
  rightRow = buildTopMxPoterySborkiRight_(ss, sh, rightRow + 2);

  sh.setRowHeight(rightRow + 1, 10);
  rightRow = buildTopFioSpisalPoterySborkiRight_(ss, sh, rightRow + 2);

  sh.setRowHeight(rightRow + 1, 10);
  rightRow = buildTopTovarPoterySborkiRight_(ss, sh, rightRow + 2);

  clearGapCols_(sh, startTablesRow, Math.max(leftRow, rightRow));
}

function buildTopTovarPoterySborkiRight_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сборки");

  const byCat = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const k = values[i][10];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const cat = stripParentheses_(k == null ? "" : String(k)).trim();
    const key = cat === "" ? "—" : cat;

    allSku.add(skuKey);
    totalSum += cost;

    if (!byCat.has(key)) byCat.set(key, { sku: new Set(), sum: 0 });
    const obj = byCat.get(key);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byCat.entries()).map(([cat, obj]) => ({
    cat,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    cat: r.cat,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  return renderTopTableRight_(sh, startRow, {
    title: "ТОП ТОВАР",
    firstLabel: "НАИМЕНОВАНИЕ",
    rows: top.map(x => [x.cat, x.skuCount, x.sum, x.pct]),
    totalSku: allSku.size,
    totalSum: totalSum,
    top1Tint: "#FEE2E2"
  });
}

function buildTopEtazhPoterySborkiRight_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сборки");

  const byEtazh = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const c = values[i][2];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const etazh = formatEtazhFromMxSborka_(stripParentheses_(c == null ? "" : String(c)));
    const key = etazh.trim() === "" ? "—" : etazh;

    allSku.add(skuKey);
    totalSum += cost;

    if (!byEtazh.has(key)) byEtazh.set(key, { sku: new Set(), sum: 0 });
    const obj = byEtazh.get(key);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byEtazh.entries()).map(([etazh, obj]) => ({
    etazh,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    etazh: r.etazh,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  return renderTopTableRight_(sh, startRow, {
    title: "ТОП ЭТАЖ",
    firstLabel: "ЭТАЖ",
    rows: top.map(x => [x.etazh, x.skuCount, x.sum, x.pct]),
    totalSku: allSku.size,
    totalSum: totalSum,
    top1Tint: "#FEE2E2"
  });
}

function buildTopRyadPoterySborkiRight_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сборки");

  const byRyad = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const c = values[i][2];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const ryad = formatRyadFromMxSborka_(stripParentheses_(c == null ? "" : String(c)));
    const key = ryad.trim() === "" ? "—" : ryad;

    allSku.add(skuKey);
    totalSum += cost;

    if (!byRyad.has(key)) byRyad.set(key, { sku: new Set(), sum: 0 });
    const obj = byRyad.get(key);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byRyad.entries()).map(([ryad, obj]) => ({
    ryad,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    ryad: r.ryad,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  return renderTopTableRight_(sh, startRow, {
    title: "ТОП РЯД",
    firstLabel: "РЯД",
    rows: top.map(x => [x.ryad, x.skuCount, x.sum, x.pct]),
    totalSku: allSku.size,
    totalSum: totalSum,
    top1Tint: "#FEE2E2"
  });
}

function buildTopMxPoterySborkiRight_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сборки");

  const byMx = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const c = values[i][2];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const mxRaw = stripParentheses_(c == null ? "" : String(c)).trim();
    const mxFmt = formatTopMxFromMxSborka_(mxRaw);
    const key = mxFmt.trim() === "" ? "—" : mxFmt;

    allSku.add(skuKey);
    totalSum += cost;

    if (!byMx.has(key)) byMx.set(key, { sku: new Set(), sum: 0 });
    const obj = byMx.get(key);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byMx.entries()).map(([mx, obj]) => ({
    mx,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    mx: r.mx,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  return renderTopTableRight_(sh, startRow, {
    title: "ТОП МХ",
    firstLabel: "МХ",
    rows: top.map(x => [x.mx, x.skuCount, x.sum, x.pct]),
    totalSku: allSku.size,
    totalSum: totalSum,
    top1Tint: "#FEE2E2"
  });
}

function buildTopFioSpisalPoterySborkiRight_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сборки");

  const byFio = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const fioE = values[i][4];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const fio = stripParentheses_(fioE == null ? "" : String(fioE)).trim();
    const key = fio === "" ? "—" : fio;

    allSku.add(skuKey);
    totalSum += cost;

    if (!byFio.has(key)) byFio.set(key, { sku: new Set(), sum: 0 });
    const obj = byFio.get(key);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byFio.entries()).map(([fio, obj]) => ({
    fio,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    fio: r.fio,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  return renderTopTableRight_(sh, startRow, {
    title: "ТОП СОТРУДНИК",
    firstLabel: "ФИО",
    rows: top.map(x => [x.fio, x.skuCount, x.sum, x.pct]),
    totalSku: allSku.size,
    totalSum: totalSum,
    top1Tint: "#FEE2E2"
  });
}

function renderTopTableRight_(sh, startRow, cfg) {
  const c0 = 8;

  const titleRow = startRow;
  const headerRow = titleRow + 1;
  const firstDataRow = headerRow + 1;

  const rows = cfg.rows || [];
  const dataCount = Math.max(1, rows.length);

  sh.setRowHeight(titleRow, 28);
  const t = sh.getRange(titleRow, c0, 1, 5);
  t.merge();
  t
    .setValue(cfg.title)
    .setFontFamily("Inter")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#EEF2FF");
  applyOuterThick_(t);

  sh.setRowHeight(headerRow, 30);

  sh.getRange(headerRow, c0, 1, 2).mergeAcross();
  sh.getRange(headerRow, c0).setValue(cfg.firstLabel);
  sh.getRange(headerRow, c0 + 2).setValue("КОЛ ВО ШК");
  sh.getRange(headerRow, c0 + 3).setValue("СУММА");
  sh.getRange(headerRow, c0 + 4).setValue("ПРОЦЕНТ");

  const h = sh.getRange(headerRow, c0, 1, 5);
  h
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(h);

  let bestIdx = 0;
  let bestSum = -1;
  for (let i = 0; i < rows.length; i++) {
    const s = Number(rows[i][2]) || 0;
    if (s > bestSum) {
      bestSum = s;
      bestIdx = i;
    }
  }

  for (let i = 0; i < dataCount; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    const item = rows.length ? rows[i] : ["—", 0, 0, 0];

    sh.getRange(row, c0, 1, 2).mergeAcross();
    sh.getRange(row, c0).setValue(item[0]);
    sh.getRange(row, c0 + 2).setValue(item[1]);
    sh.getRange(row, c0 + 3).setValue(item[2]);
    sh.getRange(row, c0 + 4).setValue(item[3]);

    const rr = sh.getRange(row, c0, 1, 5);
    rr
      .setFontFamily("Inter")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontColor("#0B1220")
      .setVerticalAlignment("middle")
      .setBackground(i % 2 ? "#F8FAFC" : "#FFFFFF");

    sh.getRange(row, c0).setHorizontalAlignment("left");
    sh.getRange(row, c0 + 2).setHorizontalAlignment("center").setNumberFormat("0");
    sh.getRange(row, c0 + 3).setHorizontalAlignment("center").setNumberFormat('#,##0" ₽"');
    sh.getRange(row, c0 + 4).setHorizontalAlignment("center").setNumberFormat("0.00%");
  }

  if (rows.length > 0) {
    sh.getRange(firstDataRow + bestIdx, c0, 1, 5).setBackground(cfg.top1Tint || "#FEE2E2");
  }

  const totalRow = firstDataRow + dataCount;
  sh.setRowHeight(totalRow, 28);

  sh.getRange(totalRow, c0, 1, 2).mergeAcross();
  sh.getRange(totalRow, c0).setValue("ИТОГО");
  sh.getRange(totalRow, c0 + 2).setValue(cfg.totalSku || 0);
  sh.getRange(totalRow, c0 + 3).setValue(cfg.totalSum || 0);
  sh.getRange(totalRow, c0 + 4).setValue(1);

  const tr = sh.getRange(totalRow, c0, 1, 5);
  tr
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(totalRow, c0 + 2).setNumberFormat("0");
  sh.getRange(totalRow, c0 + 3).setNumberFormat('#,##0" ₽"');
  sh.getRange(totalRow, c0 + 4).setNumberFormat("0.00%");

  const full = sh.getRange(titleRow, c0, totalRow - titleRow + 1, 5);
  applyProBorders_(full);
  full.setWrap(true);

  return totalRow;
}

function formatEtazhFromMxSborka_(mx) {
  const s = String(mx == null ? "" : mx).trim();
  if (s === "") return "";

  const dot = s.indexOf(".");
  if (dot === -1) {
    const base0 = stripParentheses_(s).trim();
    return base0 === "" ? "" : `${base0} ЭТАЖ`;
  }

  const left = stripParentheses_(s.slice(0, dot)).trim();
  const right = s.slice(dot + 1);

  const digitsAll = (right.match(/\d+/g) || []).join("");
  if (!digitsAll) return left ? `${left} ЭТАЖ` : "ЭТАЖ";

  let keep = "";
  if (digitsAll.length >= 2 && digitsAll[0] === "0") keep = digitsAll[1];
  else keep = digitsAll[0];

  const base = [left, keep].filter(x => x && x.trim() !== "").join(" ").trim();
  return base ? `${base} ЭТАЖ` : "ЭТАЖ";
}

/*
  ТОП РЯД: <до 1-й точки> <после 1-й точки> ЭТАЖ <после 2-й точки> РЯД
  Точки превращаем в пробелы (по сути: split('.')), всё остальное удаляем.
  Пример: "Э10.01.42 ..." -> "Э10 01 ЭТАЖ 42 РЯД"
*/
function formatRyadFromMxSborka_(mx) {
  const s0 = stripParentheses_(mx == null ? "" : String(mx));
  const s = s0.replace(/\s+/g, " ").trim();
  if (!s) return "";

  const parts = s.split(".").map(p => p.trim()).filter(p => p !== "");
  if (parts.length >= 3) {
    const out = [parts[0], parts[1], "ЭТАЖ", parts[2], "РЯД"].filter(Boolean).join(" ");
    return out.replace(/\s+/g, " ").trim();
  }
  if (parts.length === 2) {
    const out = [parts[0], parts[1], "ЭТАЖ", "РЯД"].filter(Boolean).join(" ");
    return out.replace(/\s+/g, " ").trim();
  }
  const out1 = [parts[0] || s, "РЯД"].join(" ");
  return out1.replace(/\s+/g, " ").trim();
}

/*
  ТОП МХ: <до 1-й точки> <после 1-й точки> ЭТАЖ <после 2-й точки> РЯД <после 3-й точки> СТЕЛЛАЖ
  Пример: "Э9С.04.30.43 ..." -> "Э9С 04 ЭТАЖ 30 РЯД 43 СТЕЛЛАЖ"
*/
function formatTopMxFromMxSborka_(mx) {
  const s0 = stripParentheses_(mx == null ? "" : String(mx));
  const s = s0.replace(/\s+/g, " ").trim();
  if (!s) return "";

  const parts = s.split(".").map(p => p.trim()).filter(p => p !== "");
  if (parts.length >= 4) {
    const out = [parts[0], parts[1], "ЭТАЖ", parts[2], "РЯД", parts[3], "СТЕЛЛАЖ"].filter(Boolean).join(" ");
    return out.replace(/\s+/g, " ").trim();
  }
  if (parts.length === 3) {
    const out = [parts[0], parts[1], "ЭТАЖ", parts[2], "РЯД", "СТЕЛЛАЖ"].filter(Boolean).join(" ");
    return out.replace(/\s+/g, " ").trim();
  }
  if (parts.length === 2) {
    const out = [parts[0], parts[1], "ЭТАЖ", "РЯД", "СТЕЛЛАЖ"].filter(Boolean).join(" ");
    return out.replace(/\s+/g, " ").trim();
  }
  return s;
}

function takeFirstToken_(s) {
  const x = String(s == null ? "" : s).trim();
  if (!x) return "";
  const m = x.match(/^([^\s]+)/);
  return m ? m[1] : "";
}

function buildGeneralWriteoffsBlock_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const map = new Map();
  const globalSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const name = values[i][0];
    const sku = values[i][1];
    const costRaw = values[i][13];

    if (!name || !sku) continue;

    const skuKey = String(sku);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    globalSku.add(skuKey);

    const key = stripParentheses_(String(name).trim());
    if (!map.has(key)) map.set(key, { skuSet: new Set(), sum: 0 });
    const obj = map.get(key);

    obj.skuSet.add(skuKey);
    obj.sum += cost;
    totalSum += cost;
  }

  const entries = Array.from(map.entries()).map(([name, obj]) => ({
    name,
    skuCount: obj.skuSet.size,
    sum: obj.sum
  }));
  entries.sort((a, b) => b.sum - a.sum);

  const titleRow = startRow;
  const headerRow = titleRow + 1;
  const firstDataRow = headerRow + 1;

  sh.setRowHeight(titleRow, 34);
  sh.setRowHeight(headerRow, 32);

  const blockTitle = sh.getRange(titleRow, 1, 1, 12);
  blockTitle.merge();
  blockTitle
    .setValue("ОБЩИЕ СПИСАНИЯ КОРПУСА")
    .setFontFamily("Inter")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#EEF2FF");
  applyOuterThick_(blockTitle);

  setRowLayout12_(sh, headerRow);
  setRowValues12_(sh, headerRow, ["ОПЕРАЦИЯ", "КОЛ ВО ШК", "СУММА", "ПРОЦЕНТ"]);
  formatHeaderRow12_(sh, headerRow);

  const out = [];
  if (entries.length === 0) {
    out.push({ name: "—", skuCount: 0, sum: 0, pct: 0 });
  } else {
    for (const e of entries) {
      out.push({
        name: e.name,
        skuCount: e.skuCount,
        sum: e.sum,
        pct: totalSum > 0 ? e.sum / totalSum : 0
      });
    }
  }

  for (let i = 0; i < out.length; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    setRowLayout12_(sh, row);
    setRowValues12_(sh, row, [out[i].name, out[i].skuCount, out[i].sum, out[i].pct]);
    formatDataRow12_(sh, row);

    if (i % 2 === 1) sh.getRange(row, 1, 1, 12).setBackground("#F8FAFC");
  }

  if (entries.length > 0) {
    sh.getRange(firstDataRow, 1, 1, 12).setBackground("#FEE2E2");
  }

  const totalRow = firstDataRow + out.length;
  sh.setRowHeight(totalRow, 28);

  setRowLayout12_(sh, totalRow);
  setRowValues12_(sh, totalRow, ["ИТОГО", globalSku.size, totalSum, 1]);
  formatTotalRow12_(sh, totalRow);

  const fullBlock = sh.getRange(titleRow, 1, totalRow - titleRow + 1, 12);
  applyProBorders_(fullBlock);
  fullBlock.setWrap(true);

  return totalRow;
}

function buildTwoSectionHeaders_(sh, row) {
  sh.setRowHeight(row, 40);

  const left = sh.getRange(row, 1, 1, 5);
  left.merge();
  left
    .setValue("ПОТЕРИ СОРТИРОВКИ")
    .setFontFamily("Inter")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#FFE4E6");
  applyOuterThick_(left);

  const gap = sh.getRange(row, 6, 1, 2);
  gap.setValue("").setBackground("#FFFFFF").setBorder(false, false, false, false, false, false);

  const right = sh.getRange(row, 8, 1, 5);
  right.merge();
  right
    .setValue("ПОТЕРИ СБОРКИ")
    .setFontFamily("Inter")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DCEBFF");
  applyOuterThick_(right);
}

function buildLossesByBlockTableSide_(ss, sh, startRow, cfg) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();

  const byBlock = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const d = values[i][3];
    const costRaw = values[i][13];

    if (!a || !b) continue;

    const aN = norm(a);
    if (!cfg.srcAllow.has(aN)) continue;

    const skuKey = String(b);
    const blockKey = stripParentheses_(d ? String(d).trim() : "—");
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    allSku.add(skuKey);
    totalSum += cost;

    if (!byBlock.has(blockKey)) byBlock.set(blockKey, { sku: new Set(), sum: 0 });
    const obj = byBlock.get(blockKey);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rows = Array.from(byBlock.entries()).map(([block, obj]) => ({
    block,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rows.sort((x, y) => y.sum - x.sum);

  const c0 = cfg.startCol;
  const cLabel = c0;
  const cSku = c0 + 2;
  const cSum = c0 + 3;
  const cPct = c0 + 4;

  const headerRow = startRow;
  const firstDataRow = headerRow + 1;

  sh.setRowHeight(headerRow, 30);

  sh.getRange(headerRow, cLabel, 1, 2).mergeAcross();
  sh.getRange(headerRow, cLabel).setValue("БЛОК ПОТЕРИ");
  sh.getRange(headerRow, cSku).setValue("КОЛ ВО ШК");
  sh.getRange(headerRow, cSum).setValue("СУММА");
  sh.getRange(headerRow, cPct).setValue("ПРОЦЕНТ");

  const h = sh.getRange(headerRow, c0, 1, 5);
  h
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(h);

  const out = rows.length ? rows : [{ block: "—", skuCount: 0, sum: 0 }];

  let bestIdx = 0;
  let bestSum = -1;
  for (let i = 0; i < out.length; i++) {
    if (out[i].sum > bestSum) {
      bestSum = out[i].sum;
      bestIdx = i;
    }
  }

  for (let i = 0; i < out.length; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    sh.getRange(row, cLabel, 1, 2).mergeAcross();
    sh.getRange(row, cLabel).setValue(out[i].block);
    sh.getRange(row, cSku).setValue(out[i].skuCount);
    sh.getRange(row, cSum).setValue(out[i].sum);
    sh.getRange(row, cPct).setValue(totalSum > 0 ? out[i].sum / totalSum : 0);

    const rr = sh.getRange(row, c0, 1, 5);
    rr
      .setFontFamily("Inter")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontColor("#0B1220")
      .setVerticalAlignment("middle")
      .setBackground(i % 2 ? "#F8FAFC" : "#FFFFFF");

    sh.getRange(row, cLabel).setHorizontalAlignment("left");
    sh.getRange(row, cSku).setHorizontalAlignment("center").setNumberFormat("0");
    sh.getRange(row, cSum).setHorizontalAlignment("center").setNumberFormat('#,##0" ₽"');
    sh.getRange(row, cPct).setHorizontalAlignment("center").setNumberFormat("0.00%");
  }

  if (out.length > 0) {
    sh.getRange(firstDataRow + bestIdx, c0, 1, 5).setBackground("#FEE2E2");
  }

  const totalRow = firstDataRow + out.length;
  sh.setRowHeight(totalRow, 28);

  sh.getRange(totalRow, cLabel, 1, 2).mergeAcross();
  sh.getRange(totalRow, cLabel).setValue("ИТОГО");
  sh.getRange(totalRow, cSku).setValue(allSku.size);
  sh.getRange(totalRow, cSum).setValue(totalSum);
  sh.getRange(totalRow, cPct).setValue(1);

  const t = sh.getRange(totalRow, c0, 1, 5);
  t
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(totalRow, cSku).setNumberFormat("0");
  sh.getRange(totalRow, cSum).setNumberFormat('#,##0" ₽"');
  sh.getRange(totalRow, cPct).setNumberFormat("0.00%");

  const full = sh.getRange(headerRow, c0, totalRow - headerRow + 1, 5);
  applyProBorders_(full);
  full.setWrap(true);

  return totalRow;
}

function buildSortingMxBucketsTableLeft_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сортировки");

  const tokens = makeTokens_();

  const cats = {
    "АВТОСОРТ": { sku: new Set(), sum: 0 },
    "СОРТ": { sku: new Set(), sum: 0 },
    "ПРЕДСОРТ": { sku: new Set(), sum: 0 },
    "ПОТОК": { sku: new Set(), sum: 0 },
    "ВНУТРЯНКА": { sku: new Set(), sum: 0 },
    "ПРОЧЕЕ": { sku: new Set(), sum: 0 }
  };

  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const c = values[i][2];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const cStr = c == null ? "" : String(c);
    const cN = norm(cStr);

    const containsAuto = containsAny_(cN, tokens.auto);

    let cat = "ПРОЧЕЕ";

    if (String(cStr).trim() === "") cat = "ВНУТРЯНКА";
    else if (containsAuto) cat = "АВТОСОРТ";
    else if (cN.indexOf(tokens.tokenPS) !== -1) cat = "ПРЕДСОРТ";
    else if (cN.indexOf(tokens.tokenPotok) !== -1 || cN.indexOf(tokens.tokenBuf) !== -1) cat = "ПОТОК";
    else {
      const isSort = (cN.indexOf(tokens.tokenKS) !== -1) || (cN.indexOf(tokens.tokenSort) !== -1);
      if (isSort) cat = "СОРТ";
      else cat = "ПРОЧЕЕ";
    }

    allSku.add(skuKey);
    totalSum += cost;

    cats[cat].sku.add(skuKey);
    cats[cat].sum += cost;
  }

  const order = ["АВТОСОРТ", "СОРТ", "ПРЕДСОРТ", "ПОТОК", "ВНУТРЯНКА", "ПРОЧЕЕ"];

  const rows = order.map(name => ({
    name,
    skuCount: cats[name].sku.size,
    sum: cats[name].sum,
    pct: totalSum > 0 ? (cats[name].sum / totalSum) : 0
  }));

  rows.sort((x, y) => y.sum - x.sum);

  const headerRow = startRow;
  const firstDataRow = headerRow + 1;

  sh.setRowHeight(headerRow, 30);

  sh.getRange(headerRow, 1, 1, 2).mergeAcross();
  sh.getRange(headerRow, 1).setValue("МХ");
  sh.getRange(headerRow, 3).setValue("КОЛ ВО ШК");
  sh.getRange(headerRow, 4).setValue("СУММА");
  sh.getRange(headerRow, 5).setValue("ПРОЦЕНТ");

  const h = sh.getRange(headerRow, 1, 1, 5);
  h
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(h);

  let bestIdx = 0;
  let bestSum = -1;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].sum > bestSum) {
      bestSum = rows[i].sum;
      bestIdx = i;
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    sh.getRange(row, 1, 1, 2).mergeAcross();
    sh.getRange(row, 1).setValue(rows[i].name);
    sh.getRange(row, 3).setValue(rows[i].skuCount);
    sh.getRange(row, 4).setValue(rows[i].sum);
    sh.getRange(row, 5).setValue(rows[i].pct);

    const rr = sh.getRange(row, 1, 1, 5);
    rr
      .setFontFamily("Inter")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontColor("#0B1220")
      .setVerticalAlignment("middle")
      .setBackground(i % 2 ? "#F8FAFC" : "#FFFFFF");

    sh.getRange(row, 1).setHorizontalAlignment("left");
    sh.getRange(row, 3).setHorizontalAlignment("center").setNumberFormat("0");
    sh.getRange(row, 4).setHorizontalAlignment("center").setNumberFormat('#,##0" ₽"');
    sh.getRange(row, 5).setHorizontalAlignment("center").setNumberFormat("0.00%");
  }

  if (rows.length > 0) {
    sh.getRange(firstDataRow + bestIdx, 1, 1, 5).setBackground("#FEE2E2");
  }

  const totalRow = firstDataRow + rows.length;
  sh.setRowHeight(totalRow, 28);

  sh.getRange(totalRow, 1, 1, 2).mergeAcross();
  sh.getRange(totalRow, 1).setValue("ИТОГО");
  sh.getRange(totalRow, 3).setValue(allSku.size);
  sh.getRange(totalRow, 4).setValue(totalSum);
  sh.getRange(totalRow, 5).setValue(1);

  const t = sh.getRange(totalRow, 1, 1, 5);
  t
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(totalRow, 3).setNumberFormat("0");
  sh.getRange(totalRow, 4).setNumberFormat('#,##0" ₽"');
  sh.getRange(totalRow, 5).setNumberFormat("0.00%");

  const full = sh.getRange(headerRow, 1, totalRow - headerRow + 1, 5);
  applyProBorders_(full);
  full.setWrap(true);

  return totalRow;
}

function buildTop10MxForPoterySCLeft_(ss, sh, startRow) {
  return buildTop10MxForSourceLeft_(ss, sh, startRow, {
    sourceName: "Потери СЦ",
    title: "СЦ Э9",
    titleBg: "#FFE4E6"
  });
}

function buildTop10MxForPoterySortirovkiFilteredLeft_(ss, sh, startRow, cfg) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сортировки");

  const tokens = makeTokens_();

  const byMx = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const c = values[i][2];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const cStrRaw = c == null ? "" : String(c);
    const cStrSan = stripParentheses_(cStrRaw).trim();
    const cN = norm(cStrRaw);

    const containsAuto = containsAny_(cN, tokens.auto);
    if (!cfg.predicate(cN, containsAuto)) continue;

    const mxDisplay = cfg.mxTransform ? cfg.mxTransform(cStrSan) : cStrSan;
    const mxKey = (mxDisplay === "") ? "—" : mxDisplay;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    allSku.add(skuKey);
    totalSum += cost;

    if (!byMx.has(mxKey)) byMx.set(mxKey, { sku: new Set(), sum: 0 });
    const obj = byMx.get(mxKey);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byMx.entries()).map(([mx, obj]) => ({
    mx,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    mx: r.mx,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  const titleRow = startRow;
  const headerRow = titleRow + 1;
  const firstDataRow = headerRow + 1;

  sh.setRowHeight(titleRow, 28);
  const t = sh.getRange(titleRow, 1, 1, 5);
  t.merge();
  t
    .setValue(cfg.title)
    .setFontFamily("Inter")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#EEF2FF");
  applyOuterThick_(t);

  sh.setRowHeight(headerRow, 30);

  sh.getRange(headerRow, 1, 1, 2).mergeAcross();
  sh.getRange(headerRow, 1).setValue("МХ");
  sh.getRange(headerRow, 3).setValue("КОЛ ВО ШК");
  sh.getRange(headerRow, 4).setValue("СУММА");
  sh.getRange(headerRow, 5).setValue("ПРОЦЕНТ");

  const h = sh.getRange(headerRow, 1, 1, 5);
  h
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(h);

  let bestIdx = -1;
  let bestSum = -1;
  for (let i = 0; i < top.length; i++) {
    if (top[i].sum > bestSum) {
      bestSum = top[i].sum;
      bestIdx = i;
    }
  }

  const dataCount = Math.max(1, top.length);
  for (let i = 0; i < dataCount; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    const mx = top.length ? top[i].mx : "—";
    const skuCount = top.length ? top[i].skuCount : 0;
    const sum = top.length ? top[i].sum : 0;
    const pct = top.length ? top[i].pct : 0;

    sh.getRange(row, 1, 1, 2).mergeAcross();
    sh.getRange(row, 1).setValue(mx);
    sh.getRange(row, 3).setValue(skuCount);
    sh.getRange(row, 4).setValue(sum);
    sh.getRange(row, 5).setValue(pct);

    const rr = sh.getRange(row, 1, 1, 5);
    rr
      .setFontFamily("Inter")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontColor("#0B1220")
      .setVerticalAlignment("middle")
      .setBackground(i % 2 ? "#F8FAFC" : "#FFFFFF");

    sh.getRange(row, 1).setHorizontalAlignment("left");
    sh.getRange(row, 3).setHorizontalAlignment("center").setNumberFormat("0");
    sh.getRange(row, 4).setHorizontalAlignment("center").setNumberFormat('#,##0" ₽"');
    sh.getRange(row, 5).setHorizontalAlignment("center").setNumberFormat("0.00%");
  }

  if (top.length && bestIdx >= 0) {
    sh.getRange(firstDataRow + bestIdx, 1, 1, 5).setBackground("#FEE2E2");
  }

  const totalRow = firstDataRow + dataCount;
  sh.setRowHeight(totalRow, 28);

  sh.getRange(totalRow, 1, 1, 2).mergeAcross();
  sh.getRange(totalRow, 1).setValue("ИТОГО");
  sh.getRange(totalRow, 3).setValue(allSku.size);
  sh.getRange(totalRow, 4).setValue(totalSum);
  sh.getRange(totalRow, 5).setValue(1);

  const tr = sh.getRange(totalRow, 1, 1, 5);
  tr
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(totalRow, 3).setNumberFormat("0");
  sh.getRange(totalRow, 4).setNumberFormat('#,##0" ₽"');
  sh.getRange(totalRow, 5).setNumberFormat("0.00%");

  const full = sh.getRange(titleRow, 1, totalRow - titleRow + 1, 5);
  applyProBorders_(full);
  full.setWrap(true);

  return totalRow;
}

function buildTop10FioSkladchikaForPoterySortirovkiLeft_(ss, sh, startRow) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm("Потери сортировки");

  const byFio = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const sku = values[i][1];
    const fioF = values[i][5];
    const costRaw = values[i][13];

    if (!a || !sku) continue;
    if (norm(a) !== srcNeed) continue;

    const skuKey = String(sku);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    const fioKey = stripParentheses_(fioF ? String(fioF).trim() : "").trim() || "—";

    allSku.add(skuKey);
    totalSum += cost;

    if (!byFio.has(fioKey)) byFio.set(fioKey, { sku: new Set(), sum: 0 });
    const obj = byFio.get(fioKey);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byFio.entries()).map(([fio, obj]) => ({
    fio,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    fio: r.fio,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  const titleRow = startRow;
  const headerRow = titleRow + 1;
  const firstDataRow = headerRow + 1;

  sh.setRowHeight(titleRow, 28);
  const t = sh.getRange(titleRow, 1, 1, 5);
  t.merge();
  t
    .setValue("ТОП СОТРУДНИК")
    .setFontFamily("Inter")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#EEF2FF");
  applyOuterThick_(t);

  sh.setRowHeight(headerRow, 30);

  sh.getRange(headerRow, 1, 1, 2).mergeAcross();
  sh.getRange(headerRow, 1).setValue("ФИО");
  sh.getRange(headerRow, 3).setValue("КОЛ ВО ШК");
  sh.getRange(headerRow, 4).setValue("СУММА");
  sh.getRange(headerRow, 5).setValue("ПРОЦЕНТ");

  const h = sh.getRange(headerRow, 1, 1, 5);
  h
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(h);

  let bestIdx = -1;
  let bestSum = -1;
  for (let i = 0; i < top.length; i++) {
    if (top[i].sum > bestSum) {
      bestSum = top[i].sum;
      bestIdx = i;
    }
  }

  const dataCount = Math.max(1, top.length);
  for (let i = 0; i < dataCount; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    const fio = top.length ? top[i].fio : "—";
    const skuCount = top.length ? top[i].skuCount : 0;
    const sum = top.length ? top[i].sum : 0;
    const pct = top.length ? top[i].pct : 0;

    sh.getRange(row, 1, 1, 2).mergeAcross();
    sh.getRange(row, 1).setValue(fio);
    sh.getRange(row, 3).setValue(skuCount);
    sh.getRange(row, 4).setValue(sum);
    sh.getRange(row, 5).setValue(pct);

    const rr = sh.getRange(row, 1, 1, 5);
    rr
      .setFontFamily("Inter")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontColor("#0B1220")
      .setVerticalAlignment("middle")
      .setBackground(i % 2 ? "#F8FAFC" : "#FFFFFF");

    sh.getRange(row, 1).setHorizontalAlignment("left");
    sh.getRange(row, 3).setHorizontalAlignment("center").setNumberFormat("0");
    sh.getRange(row, 4).setHorizontalAlignment("center").setNumberFormat('#,##0" ₽"');
    sh.getRange(row, 5).setHorizontalAlignment("center").setNumberFormat("0.00%");
  }

  if (top.length && bestIdx >= 0) {
    sh.getRange(firstDataRow + bestIdx, 1, 1, 5).setBackground("#FEE2E2");
  }

  const totalRow = firstDataRow + dataCount;
  sh.setRowHeight(totalRow, 28);

  sh.getRange(totalRow, 1, 1, 2).mergeAcross();
  sh.getRange(totalRow, 1).setValue("ИТОГО");
  sh.getRange(totalRow, 3).setValue(allSku.size);
  sh.getRange(totalRow, 4).setValue(totalSum);
  sh.getRange(totalRow, 5).setValue(1);

  const tr = sh.getRange(totalRow, 1, 1, 5);
  tr
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(totalRow, 3).setNumberFormat("0");
  sh.getRange(totalRow, 4).setNumberFormat('#,##0" ₽"');
  sh.getRange(totalRow, 5).setNumberFormat("0.00%");

  const full = sh.getRange(titleRow, 1, totalRow - titleRow + 1, 5);
  applyProBorders_(full);
  full.setWrap(true);

  return totalRow;
}

function buildTop10MxForSourceLeft_(ss, sh, startRow, cfg) {
  const dataSh = ss.getSheetByName("DATA");
  if (!dataSh) throw new Error('Лист "DATA" не найден');

  const lastRow = dataSh.getLastRow();
  const values = lastRow >= 2 ? dataSh.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  const norm = makeNormalizer_();
  const srcNeed = norm(cfg.sourceName);

  const byMx = new Map();
  const allSku = new Set();
  let totalSum = 0;

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0];
    const b = values[i][1];
    const c = values[i][2];
    const costRaw = values[i][13];

    if (!a || !b) continue;
    if (norm(a) !== srcNeed) continue;

    const mxRaw = c == null ? "" : String(c);
    const mxSan = stripParentheses_(mxRaw.trim());
    const mxKey = mxSan === "" ? "—" : mxSan;

    const skuKey = String(b);
    const cost = (typeof costRaw === "number") ? costRaw : (Number(costRaw) || 0);

    allSku.add(skuKey);
    totalSum += cost;

    if (!byMx.has(mxKey)) byMx.set(mxKey, { sku: new Set(), sum: 0 });
    const obj = byMx.get(mxKey);
    obj.sku.add(skuKey);
    obj.sum += cost;
  }

  const rowsAll = Array.from(byMx.entries()).map(([mx, obj]) => ({
    mx,
    skuCount: obj.sku.size,
    sum: obj.sum
  }));
  rowsAll.sort((x, y) => y.sum - x.sum);

  const top = rowsAll.slice(0, 10).map(r => ({
    mx: r.mx,
    skuCount: r.skuCount,
    sum: r.sum,
    pct: totalSum > 0 ? (r.sum / totalSum) : 0
  }));

  const titleRow = startRow;
  const headerRow = titleRow + 1;
  const firstDataRow = headerRow + 1;

  sh.setRowHeight(titleRow, 28);
  const t = sh.getRange(titleRow, 1, 1, 5);
  t.merge();
  t
    .setValue(cfg.title)
    .setFontFamily("Inter")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(cfg.titleBg || "#FFE4E6");
  applyOuterThick_(t);

  sh.setRowHeight(headerRow, 30);

  sh.getRange(headerRow, 1, 1, 2).mergeAcross();
  sh.getRange(headerRow, 1).setValue("МХ");
  sh.getRange(headerRow, 3).setValue("КОЛ ВО ШК");
  sh.getRange(headerRow, 4).setValue("СУММА");
  sh.getRange(headerRow, 5).setValue("ПРОЦЕНТ");

  const h = sh.getRange(headerRow, 1, 1, 5);
  h
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(h);

  let bestIdx = -1;
  let bestSum = -1;
  for (let i = 0; i < top.length; i++) {
    if (top[i].sum > bestSum) {
      bestSum = top[i].sum;
      bestIdx = i;
    }
  }

  const dataCount = Math.max(1, top.length);
  for (let i = 0; i < dataCount; i++) {
    const row = firstDataRow + i;
    sh.setRowHeight(row, 26);

    const mx = top.length ? top[i].mx : "—";
    const skuCount = top.length ? top[i].skuCount : 0;
    const sum = top.length ? top[i].sum : 0;
    const pct = top.length ? top[i].pct : 0;

    sh.getRange(row, 1, 1, 2).mergeAcross();
    sh.getRange(row, 1).setValue(mx);
    sh.getRange(row, 3).setValue(skuCount);
    sh.getRange(row, 4).setValue(sum);
    sh.getRange(row, 5).setValue(pct);

    const rr = sh.getRange(row, 1, 1, 5);
    rr
      .setFontFamily("Inter")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontColor("#0B1220")
      .setVerticalAlignment("middle")
      .setBackground(i % 2 ? "#F8FAFC" : "#FFFFFF");

    sh.getRange(row, 1).setHorizontalAlignment("left");
    sh.getRange(row, 3).setHorizontalAlignment("center").setNumberFormat("0");
    sh.getRange(row, 4).setHorizontalAlignment("center").setNumberFormat('#,##0" ₽"');
    sh.getRange(row, 5).setHorizontalAlignment("center").setNumberFormat("0.00%");
  }

  if (top.length && bestIdx >= 0) {
    sh.getRange(firstDataRow + bestIdx, 1, 1, 5).setBackground("#FEE2E2");
  }

  const totalRow = firstDataRow + dataCount;
  sh.setRowHeight(totalRow, 28);

  sh.getRange(totalRow, 1, 1, 2).mergeAcross();
  sh.getRange(totalRow, 1).setValue("ИТОГО");
  sh.getRange(totalRow, 3).setValue(allSku.size);
  sh.getRange(totalRow, 4).setValue(totalSum);
  sh.getRange(totalRow, 5).setValue(1);

  const tr = sh.getRange(totalRow, 1, 1, 5);
  tr
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(totalRow, 3).setNumberFormat("0");
  sh.getRange(totalRow, 4).setNumberFormat('#,##0" ₽"');
  sh.getRange(totalRow, 5).setNumberFormat("0.00%");

  const full = sh.getRange(titleRow, 1, totalRow - titleRow + 1, 5);
  applyProBorders_(full);
  full.setWrap(true);

  return totalRow;
}

function makeTokens_() {
  const norm = makeNormalizer_();
  const autoRaw = [
    "КС Э10 А-c 1",
    "КС Э10 А-c 2",
    "КС Э10 А-c 3",
    "КС Э9 А-c 1"
  ];
  return {
    auto: autoRaw.map(s => norm(s)),
    tokenKS: norm("КС"),
    tokenSort: norm("Сорт"),
    tokenPS: norm("П-С"),
    tokenPotok: norm("Поток"),
    tokenBuf: norm("БуфАвтосорт")
  };
}

function formatMxPotokToEtazh_(mx) {
  const s = String(mx == null ? "" : mx).trim();
  if (s === "") return "";

  const dot = s.indexOf(".");
  if (dot === -1) {
    const base0 = stripParentheses_(s).trim();
    return base0 === "" ? "" : `${base0} ЭТАЖ`;
  }

  const left = s.slice(0, dot).trim();
  const right = s.slice(dot + 1);

  const m = right.match(/\d+/);
  let digits = m ? m[0] : "";
  digits = digits.replace(/^0+/, "");

  const base = [stripParentheses_(left).trim(), digits].filter(x => x && x.trim() !== "").join(" ").trim();
  if (base === "") return "ЭТАЖ";
  return `${base} ЭТАЖ`;
}

function containsAny_(hay, needles) {
  for (let i = 0; i < needles.length; i++) {
    if (hay.indexOf(needles[i]) !== -1) return true;
  }
  return false;
}

function stripParentheses_(s) {
  const x = String(s == null ? "" : s);
  return x.replace(/\s*\([^)]*\)\s*/g, " ").replace(/\s+/g, " ").trim();
}

function clearGapCols_(sh, fromRow, toRow) {
  if (toRow < fromRow) return;
  const r = sh.getRange(fromRow, 6, toRow - fromRow + 1, 2);
  r.setBackground("#FFFFFF");
  r.setBorder(false, false, false, false, false, false);
}

function makeNormalizer_() {
  const map = {
    "а": "a", "А": "a",
    "в": "b", "В": "b",
    "с": "c", "С": "c",
    "е": "e", "Е": "e",
    "н": "h", "Н": "h",
    "к": "k", "К": "k",
    "м": "m", "М": "m",
    "о": "o", "О": "o",
    "р": "p", "Р": "p",
    "т": "t", "Т": "t",
    "х": "x", "Х": "x",
    "у": "y", "У": "y"
  };

  return function (s) {
    let x = String(s == null ? "" : s).toLowerCase();
    x = x.replace(/ё/g, "е");
    x = x.replace(/[–—]/g, "-");
    x = x.replace(/\s+/g, " ").trim();
    let out = "";
    for (let i = 0; i < x.length; i++) {
      const ch = x[i];
      out += (map[ch] != null) ? map[ch] : ch;
    }
    return out;
  };
}

function setRowLayout12_(sh, row) {
  sh.getRange(row, 1, 1, 6).mergeAcross();
  sh.getRange(row, 7, 1, 2).mergeAcross();
  sh.getRange(row, 9, 1, 2).mergeAcross();
  sh.getRange(row, 11, 1, 2).mergeAcross();
}

function setRowValues12_(sh, row, vals4) {
  sh.getRange(row, 1).setValue(vals4[0]);
  sh.getRange(row, 7).setValue(vals4[1]);
  sh.getRange(row, 9).setValue(vals4[2]);
  sh.getRange(row, 11).setValue(vals4[3]);
}

function formatHeaderRow12_(sh, row) {
  const r = sh.getRange(row, 1, 1, 12);
  r
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#111827");
  applyProBorders_(r);
}

function formatDataRow12_(sh, row) {
  const r = sh.getRange(row, 1, 1, 12);
  r
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setVerticalAlignment("middle")
    .setBackground("#FFFFFF");

  sh.getRange(row, 1).setHorizontalAlignment("left");
  sh.getRange(row, 7).setHorizontalAlignment("center");
  sh.getRange(row, 9).setHorizontalAlignment("center");
  sh.getRange(row, 11).setHorizontalAlignment("center");

  sh.getRange(row, 7, 1, 2).setNumberFormat("0");
  sh.getRange(row, 9, 1, 2).setNumberFormat('#,##0" ₽"');
  sh.getRange(row, 11, 1, 2).setNumberFormat("0.00%");
}

function formatTotalRow12_(sh, row) {
  const r = sh.getRange(row, 1, 1, 12);
  r
    .setFontFamily("Inter")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#0B1220")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#DBEAFE");

  sh.getRange(row, 1).setHorizontalAlignment("center");

  sh.getRange(row, 7, 1, 2).setNumberFormat("0");
  sh.getRange(row, 9, 1, 2).setNumberFormat('#,##0" ₽"');
  sh.getRange(row, 11, 1, 2).setNumberFormat("0.00%");
}

function applyProBorders_(range) {
  const black = "#000000";
  const THICK = SpreadsheetApp.BorderStyle.SOLID_THICK;
  const THIN = SpreadsheetApp.BorderStyle.SOLID;

  range.setBorder(false, false, false, false, false, false);

  range.setBorder(true, true, true, true, true, true, black, THIN);
  range.setBorder(true, true, true, true, false, false, black, THICK);
  range.setBorder(false, false, false, false, true, false, black, THICK);
}

function applyOuterThick_(range) {
  const black = "#000000";
  range.setBorder(true, true, true, true, false, false, black, SpreadsheetApp.BorderStyle.SOLID_THICK);
}
