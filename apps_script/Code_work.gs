/**
 * 보고서 작성기 online work-report calculator.
 *
 * This file is installed only in the 업무보고서 spreadsheet.
 * Trace path: raw tabs -> 계산_* tabs -> PL report tabs -> 검증/실행로그
 */

const RAW = {
  CONTRACT_ALL: '계약명세',
  INVESTOR: '기준일투자자',
  CONTRACT_SNAPSHOT: '기준일계약명세',
};

const CALC = {
  CONTRACT_ALL: '계산_계약명세',
  INVESTOR: '계산_기준일투자자',
  CONTRACT_SNAPSHOT: '계산_기준일계약명세',
  VERIFY: '검증',
  LOG: '실행로그',
};

const WORK_REPORTS = ['PL302', 'PL303', 'PL305', 'PL307', 'PL309', 'PL313', 'PL315', 'PL317', 'PL319'];
const RUN_SHEET = '실행';
const RUN_CHECKBOX_CELL = 'B7';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui
    .createMenu('보고서 작성기')
    .addItem('보고서 재계산', 'recalculateAllReportsWithNotice')
    .addSeparator()
    .addItem('실행 버튼 만들기/갱신', 'setupRunButton')
    .addItem('실행 시트 열기', 'openRunSheet')
    .addToUi();
  try {
    setupRunButton({ silent: true });
  } catch (error) {
    console.log(error);
  }
}

function onEdit(e) {
  if (!e || !e.range || e.value !== 'TRUE') return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== RUN_SHEET || e.range.getA1Notation() !== RUN_CHECKBOX_CELL) return;
  e.range.setValue(false);
  recalculateAllReportsWithNotice();
}

function recalculateAllReportsWithNotice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startedAt = new Date();
  ss.toast('보고서 재계산을 시작했습니다.', '보고서 작성기', 5);
  try {
    recalculateAllReports();
    setupRunButton({ silent: true });
    const seconds = Math.round((new Date().getTime() - startedAt.getTime()) / 1000);
    ss.toast('보고서 재계산을 완료했습니다. (' + seconds + '초)', '보고서 작성기', 10);
  } catch (error) {
    ss.toast('보고서 재계산 실패: ' + error.message, '보고서 작성기', 10);
    throw error;
  }
}

function setupRunButton(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, RUN_SHEET);
  sheet.clear();
  sheet.setFrozenRows(0);
  try {
    sheet.setHiddenGridlines(true);
  } catch (error) {
    console.log(error);
  }

  sheet.setColumnWidth(1, 36);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setRowHeight(1, 48);
  sheet.setRowHeight(4, 48);
  sheet.setRowHeight(7, 40);

  sheet.getRange('A1:D1')
    .merge()
    .setValue('보고서 작성기 보고서 실행')
    .setFontWeight('bold')
    .setFontSize(18)
    .setFontColor('#ffffff')
    .setBackground('#14213d')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C5')
    .merge()
    .setValue('보고서 재계산')
    .setFontWeight('bold')
    .setFontSize(16)
    .setFontColor('#ffffff')
    .setBackground('#1f6feb')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(RUN_CHECKBOX_CELL)
    .insertCheckboxes()
    .setValue(false)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#eaf3ff');

  sheet.getRange('C7:D7')
    .merge()
    .setValue('체크하면 보고서 재계산이 실행됩니다.')
    .setFontColor('#1f2933')
    .setVerticalAlignment('middle');

  sheet.getRange('B9:D11').setValues([
    ['실행 방법', '상태', '비고'],
    ['시트 버튼', 'B7 체크', '계산 완료 후 자동 해제'],
    ['상단 메뉴', '보고서 작성기 > 보고서 재계산', '시트를 다시 열면 메뉴 표시'],
  ]);
  sheet.getRange('B9:D9')
    .setFontWeight('bold')
    .setBackground('#d7dce3');
  sheet.getRange('B9:D11')
    .setBorder(true, true, true, true, true, true, '#9aa8b8', SpreadsheetApp.BorderStyle.SOLID);

  if (!options || !options.silent) {
    ss.setActiveSheet(sheet);
    ss.toast('실행 버튼을 준비했습니다.', '보고서 작성기', 5);
  }
}

function openRunSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, RUN_SHEET);
  ss.setActiveSheet(sheet);
}

function recalculateAllReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startedAt = new Date();
  const settings = readSettings_(ss);
  const period = readPeriod_(ss, settings, WORK_REPORTS);
  const year = period.year;
  const month = period.month;

  const rawContractAll = readTable_(ss, RAW.CONTRACT_ALL);
  const rawInvestor = readTable_(ss, RAW.INVESTOR);
  const rawContractSnapshot = readTable_(ss, RAW.CONTRACT_SNAPSHOT);

  writeReportPeriod_(ss, year, month);
  writeReadmePeriod_(ss, year, month);
  const normalized = normalizeSources_(ss, rawContractAll, rawInvestor, rawContractSnapshot, year, month);
  clearReportOutputRanges_(ss);

  recalcPL302_(ss, normalized.investor, year, month);
  recalcPL303_(ss, normalized.investor, year, month);
  recalcPL305_(ss, normalized.investor, normalized.contractAll, year, month);
  recalcPL307_(ss, normalized.contractSnapshot, year, month);
  recalcPL309_(ss, normalized.contractSnapshot, year, month);
  recalcPL313_(ss, normalized.contractSnapshot, year, month);
  recalcPL315_(ss, normalized.contractSnapshot, year, month);
  recalcPL317_(ss, normalized.contractSnapshot, year, month);
  recalcPL319_(ss, normalized.contractSnapshot, year, month);

  writeVerification_(ss, normalized, year, month);
  appendLog_(ss, startedAt, new Date(), year, month, normalized);
}

function writeReportPeriod_(ss, year, month) {
  WORK_REPORTS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.getRange(1, 1, 1, 2).setValues([[year, month]]);
  });
}

function clearReportOutputRanges_(ss) {
  workOutputRanges_().forEach(spec => clearOutputRange_(ss, spec));
}

function workOutputRanges_() {
  return [
    { sheetName: 'PL302', row: 14, col: 5, rows: 25, cols: 14 },
    { sheetName: 'PL303', row: 14, col: 5, rows: 25, cols: 14 },
    { sheetName: 'PL305', row: 8, col: 2, rows: 'max', cols: 9 },
    { sheetName: 'PL305', row: 6, col: 13, rows: 'max', cols: 12 },
    { sheetName: 'PL307', row: 14, col: 5, rows: 25, cols: 8 },
    { sheetName: 'PL309', row: 14, col: 5, rows: 25, cols: 14 },
    { sheetName: 'PL313', row: 14, col: 4, rows: 10, cols: 14 },
    { sheetName: 'PL315', row: 14, col: 5, rows: 25, cols: 12 },
    { sheetName: 'PL317', row: 15, col: 5, rows: 25, cols: 19 },
    { sheetName: 'PL319', row: 14, col: 5, rows: 25, cols: 5 },
  ];
}

function normalizeSources_(ss, contractAll, investor, contractSnapshot, year, month) {
  const contractAllNorm = normalizeContractAll_(contractAll);
  const contractByNo = {};
  contractAllNorm.forEach(row => {
    const no = text_(row['계약No']);
    if (no) contractByNo[no] = row;
  });

  const investorNorm = normalizeInvestor_(investor, contractByNo, year, month);
  const snapshotNorm = normalizeContractSnapshot_(contractSnapshot, contractByNo, year, month);

  writeObjects_(ss, CALC.CONTRACT_ALL, contractAllNorm);
  writeObjects_(ss, CALC.INVESTOR, investorNorm);
  writeObjects_(ss, CALC.CONTRACT_SNAPSHOT, snapshotNorm);

  return {
    contractAll: contractAllNorm,
    investor: investorNorm,
    contractSnapshot: snapshotNorm,
  };
}

function normalizeContractAll_(rows) {
  return rows.map(row => {
    const out = clone_(row);
    applyAdjustedStatus_(out);
    out['상품명_정리'] = cleanProductName_(out['상품명']);
    return out;
  });
}

function normalizeInvestor_(rows, contractByNo, year, month) {
  return rows.map(row => {
    const out = clone_(row);
    out['기준년'] = valueOr_(out['기준년'], year);
    out['기준월'] = valueOr_(out['기준월'], month);
    const contract = contractByNo[text_(out['계약번호'])] || {};
    out['대출자구분'] = valueOr_(out['대출자구분'], contract['대출자구분']);
    out['투자자구분_정리'] = normalizeInvestorType_(out['투자자구분']);
    out['상품분류_정리'] = normalizeProduct_(out['세부상품'] || contract['세부상품']);
    out['투자금액_숫자'] = number_(out['투자금액']);
    out['투자잔액_숫자'] = number_(out['투자잔액']);
    return out;
  });
}

function writeVerification_(ss, normalized, year, month) {
  const ym = period_(year, month);
  const prev = addPeriod_(ym, -1);
  const currentInvestor = normalized.investor.filter(row => samePeriod_(row, year, month));
  const currentSnapshot = normalized.contractSnapshot.filter(row => samePeriod_(row, year, month));
  const prevSnapshot = normalized.contractSnapshot.filter(row => samePeriod_(row, prev.year, prev.month));
  const snapshotPeriods = Array.from(new Set(normalized.contractSnapshot
    .map(row => [Number(row['기준년']), Number(row['기준월'])])
    .filter(pair => pair[0] && pair[1])
    .map(pair => pair[0] + '-' + String(pair[1]).padStart(2, '0'))))
    .sort();
  const rows = [
    ['검증항목', '값', '비고'],
    ['기준년월', year + '-' + month, 'README 기준'],
    ['계약명세 원천 행수', normalized.contractAll.length, RAW.CONTRACT_ALL],
    ['기준일투자자 원천 행수', normalized.investor.length, RAW.INVESTOR],
    ['기준일계약명세 원천 행수', normalized.contractSnapshot.length, RAW.CONTRACT_SNAPSHOT],
    ['당월 투자자 행수', currentInvestor.length, '계산_기준일투자자'],
    ['당월 기준일계약 행수', currentSnapshot.length, '계산_기준일계약명세'],
    ['전월 기준일계약 행수', prevSnapshot.length, 'PL307 전월 잔액/상환 계산 원천'],
    ['기준일계약 보유 기간', snapshotPeriods.join(', '), '전월 데이터가 없으면 PL307은 엑셀 이력 계산과 일치할 수 없습니다.'],
    ['당월 투자잔액 합계', sum_(currentInvestor, '투자잔액_숫자'), 'PL302/PL303 원천'],
    ['당월 대출잔액 합계', sum_(currentSnapshot, '잔액_숫자'), 'PL307/309/315/317/319 원천'],
    ['연체 대출잔액 합계', sum_(currentSnapshot.filter(row => text_(row['상태']) === '연체'), '잔액_숫자'), 'PL317 원천'],
  ];
  writeValues_(ss, CALC.VERIFY, rows);
}

function normalizeContractSnapshot_(rows, contractByNo, year, month) {
  const prepared = rows.map(row => {
    const out = clone_(row);
    const contract = contractByNo[text_(out['계약No'])] || {};
    applyAdjustedStatus_(out, contract);
    out['기준년'] = valueOr_(out['기준년'], year);
    out['기준월'] = valueOr_(out['기준월'], month);
    out['잔액_숫자'] = number_(out['잔액']);
    out['총대출액_숫자'] = number_(out['총대출액']);
    out['LTV_숫자'] = number_(out['LTV']);
    out['기간_숫자'] = number_(out['기간']);
    out['월-계약No'] = [out['기준년'], out['기준월'], out['계약No']].map(text_).join('-');
    out['월-계약No-상태'] = [out['기준년'], out['기준월'], out['계약No'], out['상태']].map(text_).join('-');
    out['월-주민등록번호'] = [out['기준년'], out['기준월'], out['주민등록']].map(text_).join('-');
    out['상품분류_정리'] = normalizeProduct_(out['세부상품']);
    out['PL309_계산'] = valueOr_(out['PL309'], amountBand_(out['잔액_숫자']));
    out['PL313_계산'] = valueOr_(out['PL313'], ltvBand_(out['LTV_숫자']));
    out['PL315_계산'] = valueOr_(out['PL315'], termBand_(out['기간_숫자']));
    out['PL317_계산'] = valueOr_(out['PL317'], overdueBand_(out));
    return out;
  });

  const byPersonMonth = groupBy_(prepared, row => [row['기준년'], row['기준월'], row['주민등록']].map(text_).join('|'));
  const byContractMonth = groupBy_(prepared, row => [row['기준년'], row['기준월'], row['계약No']].map(text_).join('|'));

  prepared.forEach(row => {
    const ym = period_(Number(row['기준년']), Number(row['기준월']));
    const prev = addPeriod_(ym, -1);
    const next = addPeriod_(ym, 1);
    const person = text_(row['주민등록']);
    const currentRows = byPersonMonth[[ym.year, ym.month, person].join('|')] || [];
    const prevRows = byPersonMonth[[prev.year, prev.month, person].join('|')] || [];
    const nextRows = byPersonMonth[[next.year, next.month, person].join('|')] || [];

    const maxBalance = max_(currentRows.map(r => r['잔액_숫자']));
    const maxLtv = max_(currentRows.map(r => r['LTV_숫자']));
    const maxTerm = max_(currentRows.map(r => r['기간_숫자']));
    const maxContractNo = max_(currentRows.map(r => number_(r['계약No'])));
    const maxBalanceContractNo = maxContractForValue_(currentRows, '잔액_숫자', maxBalance);
    const maxLtvContractNo = maxContractForValue_(currentRows, 'LTV_숫자', maxLtv);
    const maxTermContractNo = maxContractForValue_(currentRows, '기간_숫자', maxTerm);

    row['pl307 신규차입자'] = sum_(prevRows, '잔액_숫자') === 0 && number_(row['계약No']) === maxContractNo ? 1 : 0;
    const prevContractRows = byContractMonth[[prev.year, prev.month, row['계약No']].map(text_).join('|')] || [];
    row['pl307 신규취급액'] = prevContractRows.length ? 0 : Math.max(0, number_(row['총대출액_숫자']));

    row['pl309 중복확인'] = row['잔액_숫자'] !== 0 && row['잔액_숫자'] === maxBalance && number_(row['계약No']) === maxBalanceContractNo ? 1 : 0;
    row['pl313 중복확인'] = row['잔액_숫자'] !== 0 && row['LTV_숫자'] === maxLtv && number_(row['계약No']) === maxLtvContractNo ? 1 : 0;
    row['pl315 중복확인'] = row['잔액_숫자'] !== 0 && row['기간_숫자'] === maxTerm && number_(row['계약No']) === maxTermContractNo ? 1 : 0;

    const repaymentPeriod = row['잔액_숫자'] === 0 ? prev : ym;
    row['pl307 상환기준년'] = repaymentPeriod.year;
    row['pl307 상환기준월'] = repaymentPeriod.month;
    row['pl307 완제차입자'] = nextRows.length ? (row['잔액_숫자'] === 0 ? 1 : 0) : (number_(row['계약No']) === maxContractNo ? 1 : 0);
    const nextContractRows = byContractMonth[[next.year, next.month, row['계약No']].map(text_).join('|')] || [];
    const nextBalance = nextContractRows.length ? number_(nextContractRows[0]['잔액_숫자']) : 0;
    row['pl307 상환액'] = row['잔액_숫자'] - nextBalance + (row['잔액_숫자'] === 0 ? row['총대출액_숫자'] : 0);
  });

  return prepared;
}

function applyAdjustedStatus_(row, contract) {
  const rawStatus = text_(row['상태']);
  const overdueDays = overdueDays_(row, contract || {});
  const adjusted = adjustedStatus_(rawStatus, overdueDays);
  row['상태_원천'] = rawStatus;
  row['연체일_숫자'] = overdueDays === null ? '' : overdueDays;
  row['상태분류_계산'] = adjusted;
  row['상태'] = adjusted;
}

function adjustedStatus_(rawStatus, overdueDays) {
  const status = text_(rawStatus);
  if (status.includes('매각') || status.includes('완제') || status.includes('완납')) return status;
  if (overdueDays !== null) return overdueDays >= 30 ? '연체' : '정상';
  if (status.includes('연체')) return '연체';
  return status || '정상';
}

function overdueDays_(row, contract) {
  const rowValue = firstPresent_(row, ['연체일', '최장연체일수']);
  if (rowValue !== null) return number_(rowValue);
  const contractValue = firstPresent_(contract || {}, ['연체일', '최장연체일수']);
  if (contractValue !== null) return number_(contractValue);
  return null;
}

function firstPresent_(row, keys) {
  for (let i = 0; i < keys.length; i++) {
    const value = row[keys[i]];
    if (value !== null && value !== undefined && value !== '') return value;
  }
  return null;
}

function recalcPL302_(ss, investor, year, month) {
  const rows = productRows_().map(def => {
    const filtered = investor.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const result = [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || ''];
    investorTypes_().forEach(type => {
      const typed = filtered.filter(row => row['투자자구분_정리'] === type.key);
      result.push(uniqueCountNonzero_(typed, '주민등록번호'));
      result.push(sum_(typed, '투자잔액_숫자'));
    });
    result.push(sumEveryOther_(result, 4));
    result.push(sumEveryOther_(result, 5));
    return result;
  });
  writeProductReportValues_(ss, 'PL302', 14, 5, rows, 4);
}

function recalcPL303_(ss, investor, year, month) {
  const bands = investAmountBands_();
  const currentRows = investor.filter(row => samePeriod_(row, year, month));
  const periodSums = groupSum_(currentRows, '주민등록번호', '투자잔액_숫자');
  const rows = productRows_().map(def => {
    const filtered = investor.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const ids = uniqueValues_(filtered.map(row => row['주민등록번호']));
    const amountsForIds = ids.map(id => periodSums[id] || 0);
    const result = [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || ''];
    bands.forEach(band => {
      const amounts = amountsForIds.filter(value => value > band.min && value <= band.max);
      result.push(amounts.length);
      result.push(amounts.reduce((a, b) => a + b, 0));
    });
    result.push(amountsForIds.length);
    result.push(amountsForIds.reduce((a, b) => a + b, 0));
    return result;
  });
  writeProductReportValues_(ss, 'PL303', 14, 5, rows, 4);
}

function recalcPL305_(ss, investor, contractAll, year, month) {
  const contractByProduct = {};
  contractAll.forEach(row => {
    const key = cleanProductName_(row['상품명']);
    if (key) contractByProduct[key] = row;
  });
  const filtered = investor
    .filter(row => samePeriod_(row, year, month))
    .filter(row => ['법인', '전문'].includes(row['투자자구분_정리']));

  const rows = filtered.map(row => {
    const product = cleanProductName_(row['상품명']);
    const contract = contractByProduct[product] || {};
    const type = row['투자자구분_정리'] === '법인' ? '법인투자자' : '개인전문투자자';
    return [
      product,
      number_(contract['한도'] || contract['총대출'] || 0),
      row['세부상품'] || contract['세부상품'] || '',
      yyyymmdd_(contract['계약일']),
      yyyymmdd_(contract['만기일']),
      type === '개인전문투자자' ? maskName_(row['이름']) : row['이름'],
      type,
      type === '법인투자자' ? businessNo_(row['주민등록번호']) : '',
      number_(row['투자금액_숫자']),
    ];
  });
  writeRows_(ss, 'PL305', 8, 2, rows, 9);
  writePL305InvestorFilter_(ss, filtered);
}

function writePL305InvestorFilter_(ss, rows) {
  const sheet = ensureSheet_(ss, 'PL305');
  const headers = ['주민등록번호', '상품명', '세부상품', '이름', '투자자구분', '투자금액'];
  const clearRows = Math.max(sheet.getMaxRows() - 6 + 1, 1);
  sheet.getRange(6, 13, clearRows, 12).clearContent();
  sheet.getRange(6, 13, 1, headers.length).setValues([[
    '투자자리스트 필터링',
    '',
    '',
    '',
    '',
    '',
  ]]);
  sheet.getRange(6, 18).setValue(sum_(rows, '투자금액_숫자'));
  sheet.getRange(7, 13, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(8, 13, rows.length, headers.length).setValues(rows.map(row => [
      row['주민등록번호'] || '',
      cleanProductName_(row['상품명']),
      row['세부상품'] || '',
      row['이름'] || '',
      row['투자자구분'] || '',
      number_(row['투자금액_숫자']),
    ]));
  }
}

function recalcPL307_(ss, snapshots, year, month) {
  const ym = period_(year, month);
  const prev = addPeriod_(ym, -1);
  const hasPrevPeriod = snapshots.some(row => samePeriod_(row, prev.year, prev.month));
  const rows = productRows_().map(def => {
    const prevRows = hasPrevPeriod ? snapshots.filter(row => samePeriod_(row, prev.year, prev.month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products)) : [];
    const currentRows = snapshots.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const repayRows = hasPrevPeriod ? snapshots.filter(row => Number(row['pl307 상환기준년']) === prev.year && Number(row['pl307 상환기준월']) === prev.month && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products)) : [];
    const newBorrowerRows = hasPrevPeriod ? currentRows.filter(row => number_(row['pl307 신규차입자']) > 0) : [];
    const completedRows = hasPrevPeriod ? repayRows.filter(row => number_(row['pl307 완제차입자']) > 0) : [];
    const prevCount = prevRows.filter(row => number_(row['pl315 중복확인']) > 0).length;
    const prevAmount = sum_(prevRows, '잔액_숫자');
    const newCount = newBorrowerRows.length;
    const newAmount = sum_(currentRows, 'pl307 신규취급액');
    const repayCount = completedRows.length;
    const repayAmount = sum_(repayRows, 'pl307 상환액');
    const currentCount = hasPrevPeriod
      ? prevCount + newCount - repayCount
      : uniqueCount_(currentRows.filter(row => number_(row['잔액_숫자']) !== 0), '주민등록');
    const currentAmount = hasPrevPeriod
      ? prevAmount + newAmount - repayAmount
      : sum_(currentRows, '잔액_숫자');
    return [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || '', prevCount, prevAmount, newCount, newAmount, repayCount, repayAmount, currentCount, currentAmount];
  });
  writeProductReportValues_(ss, 'PL307', 14, 5, rows, 4);
}

function recalcPL309_(ss, snapshots, year, month) {
  const bands = loanAmountBands_();
  const rows = productRows_().map(def => {
    const filtered = snapshots.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const result = [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || ''];
    bands.forEach(band => {
      const bandRows = filtered.filter(row => (row['PL309_계산'] || amountBand_(row['잔액_숫자'])) === band.label);
      result.push(bandRows.filter(row => number_(row['pl309 중복확인']) > 0).length);
      result.push(sum_(bandRows, '잔액_숫자'));
    });
    result.push(sumEveryOther_(result, 4));
    result.push(sumEveryOther_(result, 5));
    return result;
  });
  writeProductReportValues_(ss, 'PL309', 14, 5, rows, 4);
}

function recalcPL313_(ss, snapshots, year, month) {
  const bands = ltvBands_();
  const defs = realEstateRows_();
  const rows = defs.map(def => {
    const filtered = snapshots.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const result = [def.code, def.borrowerLabel, def.productLabel];
    bands.forEach(band => {
      const bandRows = filtered.filter(row => (row['PL313_계산'] || ltvBand_(row['LTV_숫자'])) === band.label);
      result.push(bandRows.filter(row => number_(row['pl313 중복확인']) > 0).length);
      result.push(sum_(bandRows, '잔액_숫자'));
    });
    result.push(sumEveryOther_(result, 3));
    result.push(sumEveryOther_(result, 4));
    return result;
  });
  writeRealEstateReportValues_(ss, 'PL313', rows, 3);
}

function recalcPL315_(ss, snapshots, year, month) {
  const bands = termBands_();
  const rows = productRows_().map(def => {
    const filtered = snapshots.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const result = [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || ''];
    bands.forEach(band => {
      const bandRows = filtered.filter(row => (row['PL315_계산'] || termBand_(row['기간_숫자'])) === band.label);
      result.push(bandRows.filter(row => number_(row['pl315 중복확인']) > 0).length);
      result.push(sum_(bandRows, '잔액_숫자'));
    });
    result.push(sumEveryOther_(result, 4));
    result.push(sumEveryOther_(result, 5));
    return result;
  });
  writeProductReportValues_(ss, 'PL315', 14, 5, rows, 4);
}

function recalcPL317_(ss, snapshots, year, month) {
  const bands = termBands_();
  const rows = productRows_().map(def => {
    const filtered = snapshots.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const normalRows = filtered.filter(row => !text_(row['PL317_계산'] || overdueBand_(row)));
    const overdueRows = filtered.filter(row => text_(row['PL317_계산'] || overdueBand_(row)));
    const result = [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || ''];
    result.push(normalRows.filter(row => number_(row['pl315 중복확인']) > 0).length);
    result.push(sum_(normalRows, '잔액_숫자'));
    bands.forEach(band => {
      const bandRows = overdueRows.filter(row => (row['PL317_계산'] || overdueBand_(row)) === band.label);
      result.push(bandRows.filter(row => number_(row['pl315 중복확인']) > 0).length);
      result.push(sum_(bandRows, '잔액_숫자'));
    });
    const overdueCount = result[6] + result[8] + result[10] + result[12] + result[14];
    const overdueAmount = result[7] + result[9] + result[11] + result[13] + result[15];
    const totalCount = result[4] + overdueCount;
    const totalAmount = result[5] + overdueAmount;
    result.push(overdueCount, overdueAmount, overdueCount, overdueAmount, totalCount, totalAmount, totalAmount ? overdueAmount / totalAmount * 100 : 0);
    return result;
  });
  writePL317Values_(ss, rows);
}

function recalcPL319_(ss, snapshots, year, month) {
  const rows = productRows_().map(def => {
    const filtered = snapshots.filter(row => samePeriod_(row, year, month) && matchBorrower_(row, def.borrower) && matchProduct_(row, def.products));
    const balance = sum_(filtered, '잔액_숫자');
    const interest = filtered.reduce((total, row) => total + number_(row['잔액_숫자']) * number_(row['이율']) / 100, 0) / 12;
    const rates = filtered.map(row => number_(row['이율'])).filter(v => v > 0);
    return [def.code, def.borrowerLabel, def.productLabel, def.detailLabel || '', balance, interest, balance ? interest * 12 / balance * 100 : 0, rates.length ? Math.max(...rates) : 0, rates.length ? Math.min(...rates) : 0];
  });
  writePL319Values_(ss, rows);
}

function writeProductReportValues_(ss, sheetName, firstRow, startCol, rows, metricStartIndex) {
  const metrics = productReportMetrics_(rows, metricStartIndex);
  writeProductMetrics_(ss, sheetName, firstRow, startCol, metrics);
}

function writeProductMetrics_(ss, sheetName, firstRow, startCol, metrics) {
  const sheet = ensureSheet_(ss, sheetName);
  const order = productReportCodeOrder_();
  const width = Math.max(...Object.values(metrics).map(values => values.length));
  const values = order.map(code => pad_(metrics[code] || [], width));
  sheet.getRange(firstRow, startCol, values.length, width).clearContent();
  sheet.getRange(firstRow, startCol, values.length, width).setValues(values);
}

function productReportMetrics_(rows, metricStartIndex) {
  const out = {};
  rows.forEach(row => {
    out[text_(row[0])] = row.slice(metricStartIndex).map(number_);
  });
  const sumCodes = codes => {
    const width = Math.max(...codes.map(code => (out[code] || []).length), 0);
    const values = Array(width).fill(0);
    codes.forEach(code => {
      const row = out[code] || [];
      for (let i = 0; i < width; i++) values[i] += number_(row[i]);
    });
    return values;
  };
  ['A', 'B', 'C'].forEach(prefix => {
    out[prefix + '23'] = sumCodes([prefix + '21', prefix + '22']);
    out[prefix + '6'] = sumCodes([prefix + '1', prefix + '23', prefix + '3', prefix + '4', prefix + '5']);
  });
  out['D'] = sumCodes(['A6', 'B6', 'C6']);
  return out;
}

function writePL317Values_(ss, rows) {
  const metrics = productReportMetrics_(rows, 4);
  Object.keys(metrics).forEach(code => {
    const values = metrics[code];
    values[18] = number_(values[17]) ? number_(values[15]) / number_(values[17]) * 100 : 0;
  });
  writeProductMetrics_(ss, 'PL317', 15, 5, metrics);
}

function writePL319Values_(ss, rows) {
  const metrics = {};
  rows.forEach(row => metrics[text_(row[0])] = row.slice(4).map(number_));
  const aggregate = codes => {
    const balance = codes.reduce((total, code) => total + number_((metrics[code] || [])[0]), 0);
    const interest = codes.reduce((total, code) => total + number_((metrics[code] || [])[1]), 0);
    const maxRates = codes.map(code => number_((metrics[code] || [])[3])).filter(value => value > 0);
    const minRates = codes.map(code => number_((metrics[code] || [])[4])).filter(value => value > 0);
    return [
      balance,
      interest,
      balance ? interest * 12 / balance * 100 : 0,
      maxRates.length ? Math.max(...maxRates) : 0,
      minRates.length ? Math.min(...minRates) : 0,
    ];
  };
  ['A', 'B', 'C'].forEach(prefix => {
    metrics[prefix + '23'] = aggregate([prefix + '21', prefix + '22']);
    metrics[prefix + '6'] = aggregate([prefix + '1', prefix + '23', prefix + '3', prefix + '4', prefix + '5']);
  });
  metrics['D'] = aggregate(['A6', 'B6', 'C6']);
  writeProductMetrics_(ss, 'PL319', 14, 5, metrics);
}

function productReportCodeOrder_() {
  const order = [];
  ['A', 'B', 'C'].forEach(prefix => {
    order.push(prefix + '1', prefix + '21', prefix + '22', prefix + '23', prefix + '3', prefix + '4', prefix + '5', prefix + '6');
  });
  order.push('D');
  return order;
}

function writeRealEstateReportValues_(ss, sheetName, rows, metricStartIndex) {
  const sheet = ensureSheet_(ss, sheetName);
  const out = {};
  rows.forEach(row => {
    out[text_(row[0])] = row.slice(metricStartIndex).map(number_);
  });
  const sumCodes = codes => {
    const width = Math.max(...codes.map(code => (out[code] || []).length), 0);
    const values = Array(width).fill(0);
    codes.forEach(code => {
      const row = out[code] || [];
      for (let i = 0; i < width; i++) values[i] += number_(row[i]);
    });
    return values;
  };
  out['C'] = sumCodes(['A', 'B']);
  out['F'] = sumCodes(['D', 'E']);
  out['I'] = sumCodes(['G', 'H']);
  out['J'] = sumCodes(['C', 'F', 'I']);
  const order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
  const width = Math.max(...Object.values(out).map(values => values.length));
  const values = order.map(code => pad_(out[code] || [], width));
  sheet.getRange(14, 4, values.length, width).clearContent();
  sheet.getRange(14, 4, values.length, width).setValues(values);
}

function pad_(values, width) {
  const out = values.slice();
  while (out.length < width) out.push('');
  return out;
}

function clearOutputRange_(ss, spec) {
  const sheet = ss.getSheetByName(spec.sheetName);
  if (!sheet) return;
  const rowCount = spec.rows === 'max'
    ? Math.max(sheet.getMaxRows() - spec.row + 1, 1)
    : spec.rows;
  if (rowCount > 0 && spec.cols > 0) {
    sheet.getRange(spec.row, spec.col, rowCount, spec.cols).clearContent();
  }
}

function appendLog_(ss, startedAt, endedAt, year, month, normalized) {
  const sheet = ensureSheet_(ss, CALC.LOG);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['시작', '종료', '기준년', '기준월', '계약명세', '기준일투자자', '기준일계약명세']);
  }
  sheet.appendRow([startedAt, endedAt, year, month, normalized.contractAll.length, normalized.investor.length, normalized.contractSnapshot.length]);
}

function readSettings_(ss) {
  const out = {};
  readSettingPairs_(ss.getSheetByName('README'), out);
  return out;
}

function readSettingPairs_(sheet, out) {
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  values.forEach(row => {
    const key = text_(row[0]).trim();
    if (key && key !== '항목') out[key] = row[1];
  });
}

function readPeriod_(ss, settings, reportNames) {
  const now = new Date();
  const reportPeriod = readReportPeriod_(ss, reportNames);
  const fallbackYear = validPeriodNumber_(reportPeriod.year, now.getFullYear(), 2000, 2100);
  const fallbackMonth = validPeriodNumber_(reportPeriod.month, now.getMonth() + 1, 1, 12);
  return {
    year: validPeriodNumber_(settings['기준년'] || settings['default_year'], fallbackYear, 2000, 2100),
    month: validPeriodNumber_(settings['기준월'] || settings['default_month'], fallbackMonth, 1, 12),
  };
}

function readReportPeriod_(ss, reportNames) {
  for (const name of reportNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const row = sheet.getRange(1, 1, 1, 2).getValues()[0];
    return { year: row[0], month: row[1] };
  }
  return { year: '', month: '' };
}

function validPeriodNumber_(value, fallback, min, max) {
  const parsed = parsePeriodNumber_(value);
  if (parsed !== null && parsed >= min && parsed <= max) return parsed;
  return fallback;
}

function parsePeriodNumber_(value) {
  if (typeof value === 'number' && isFinite(value)) return Math.round(value);
  if (value instanceof Date && !isNaN(value.getTime())) return spreadsheetDateSerial_(value);
  const text = text_(value).trim();
  if (!text || text.charAt(0) === '#') return null;
  const dateMatch = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (dateMatch) {
    return spreadsheetDateSerial_(new Date(Number(dateMatch[1]), Number(dateMatch[2]) - 1, Number(dateMatch[3])));
  }
  const direct = Number(text.replace(/,/g, ''));
  if (isFinite(direct)) return Math.round(direct);
  const firstNumber = text.match(/-?\d+(\.\d+)?/);
  if (firstNumber) return Math.round(Number(firstNumber[0]));
  return null;
}

function spreadsheetDateSerial_(date) {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const base = Date.UTC(1899, 11, 30);
  return Math.round((utc - base) / 86400000);
}

function writeReadmePeriod_(ss, year, month) {
  const sheet = ss.getSheetByName('README');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const updates = { '기준년': String(year), '기준월': String(month) };
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 0; i < values.length; i++) {
      if (text_(values[i][0]).trim() === key) {
        sheet.getRange(i + 1, 2).setNumberFormat('@').setValue(updates[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      const row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 2).setValues([[key, updates[key]]]);
      sheet.getRange(row, 2).setNumberFormat('@');
    }
  });
}

function readTable_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(v => text_(v).trim());
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const item = {};
    headers.forEach((header, i) => {
      if (header) item[header] = row[i];
    });
    return item;
  });
}

function writeObjects_(ss, sheetName, objects) {
  if (!objects.length) {
    writeValues_(ss, sheetName, [['데이터 없음']]);
    return;
  }
  ensureSheet_(ss, sheetName).clearContents();
  const headers = Object.keys(objects[0]);
  writeTable_(ss, sheetName, 1, 1, headers, objects.map(row => headers.map(header => row[header])));
}

function writeTable_(ss, sheetName, startRow, startCol, headers, rows) {
  const sheet = ensureSheet_(ss, sheetName);
  const width = headers.length;
  const clearRows = Math.max(sheet.getMaxRows() - startRow + 1, 1);
  sheet.getRange(startRow, startCol, clearRows, Math.max(width, 1)).clearContent();
  sheet.getRange(startRow, startCol, 1, width).setValues([headers]);
  if (rows.length) {
    sheet.getRange(startRow + 1, startCol, rows.length, width).setValues(rows);
  }
  if (startRow === 1) sheet.setFrozenRows(1);
}

function writeRows_(ss, sheetName, startRow, startCol, rows, width) {
  const sheet = ensureSheet_(ss, sheetName);
  const clearRows = Math.max(sheet.getMaxRows() - startRow + 1, 1);
  sheet.getRange(startRow, startCol, clearRows, Math.max(width, 1)).clearContent();
  if (rows.length) {
    sheet.getRange(startRow, startCol, rows.length, width).setValues(rows.map(row => pad_(row, width)));
  }
}

function writeValues_(ss, sheetName, rows) {
  const sheet = ensureSheet_(ss, sheetName);
  sheet.clearContents();
  if (rows.length) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}

function ensureSheet_(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function productRows_() {
  const borrowers = [
    ['A', '개인', '개인대출'],
    ['B', '개인사업자', '개인 사업자 대출'],
    ['C', '법인', '법인대출'],
  ];
  const products = [
    ['1', ['신용대출'], '신용대출', ''],
    ['21', ['부동산담보대출'], '부동산담보대출', '주택담보대출'],
    ['22', ['기타부동산 담보대출'], '부동산담보대출', '기타담보대출'],
    ['3', ['부동산 PF대출'], '부동산PF', ''],
    ['4', ['동산 담보대출', '기타담보대출', '임차보증금 담보대출', 'NPL채권 담보대출', 'ABL 대출', '유가증권 담보대출'], '기타담보(어음, 매출채권 제외)', ''],
    ['5', ['매출채권 담보대출'], '어음, 매출채권 담보', ''],
  ];
  const rows = [];
  borrowers.forEach(([prefix, borrower, borrowerLabel]) => {
    products.forEach(([code, productList, productLabel, detailLabel]) => {
      rows.push({ code: prefix + code, borrower, borrowerLabel, productLabel, detailLabel, products: productList });
    });
  });
  return rows;
}

function realEstateRows_() {
  return [
    { code: 'A', borrower: '개인', borrowerLabel: '개인 대출', productLabel: '주택담보대출', products: ['부동산담보대출'] },
    { code: 'B', borrower: '개인', borrowerLabel: '개인 대출', productLabel: '기타담보대출', products: ['기타부동산 담보대출'] },
    { code: 'D', borrower: '개인사업자', borrowerLabel: '개인 사업자 대출', productLabel: '주택담보대출', products: ['부동산담보대출'] },
    { code: 'E', borrower: '개인사업자', borrowerLabel: '개인 사업자 대출', productLabel: '기타담보대출', products: ['기타부동산 담보대출'] },
    { code: 'G', borrower: '법인', borrowerLabel: '법인 대출', productLabel: '주택담보대출', products: ['부동산담보대출'] },
    { code: 'H', borrower: '법인', borrowerLabel: '법인 대출', productLabel: '기타담보대출', products: ['기타부동산 담보대출'] },
  ];
}

function investorTypes_() {
  return [
    { key: '개인', label: '일반투자자' },
    { key: '소득적격', label: '소득적격 개인투자자' },
    { key: '전문', label: '개인전문투자자' },
    { key: '법인', label: '일반법인' },
    { key: '금융기관', label: '금융기관' },
    { key: '온투업자', label: '온라인투자연계금융업자' },
  ];
}

function investAmountBands_() {
  return [
    { label: '1백만원 이하', min: -Infinity, max: 1000000 },
    { label: '1백만원 초과 5백만원 이하', min: 1000000, max: 5000000 },
    { label: '5백만원 초과 1천만원이하', min: 5000000, max: 10000000 },
    { label: '1천만원 초과 3천만원이하', min: 10000000, max: 30000000 },
    { label: '3천만원 초과 5천만원이하', min: 30000000, max: 50000000 },
    { label: '5천만원 초과', min: 50000000, max: Infinity },
  ];
}

function loanAmountBands_() {
  return [
    { label: '5천만원 이하', min: -Infinity, max: 50000000 },
    { label: '5천만원 초과 1억원이하', min: 50000000, max: 100000000 },
    { label: '1억원 초과 5억원 이하', min: 100000000, max: 500000000 },
    { label: '5억원 초과 30억원이하', min: 500000000, max: 3000000000 },
    { label: '30억원 초과 70억원이하', min: 3000000000, max: 7000000000 },
    { label: '70억원 초과', min: 7000000000, max: Infinity },
  ];
}

function ltvBands_() {
  return [
    { label: '40% 이하', min: -Infinity, max: 40 },
    { label: '40% 초과 50%이하', min: 40, max: 50 },
    { label: '50% 초과 60%이하', min: 50, max: 60 },
    { label: '60% 초과 70%이하', min: 60, max: 70 },
    { label: '70% 초과 100%이하', min: 70, max: 100 },
    { label: '100% 초과', min: 100, max: Infinity },
  ];
}

function termBands_() {
  return [
    { label: '3개월 이하', min: -Infinity, max: 3 },
    { label: '3개월 초과 6개월 이하', min: 3, max: 6 },
    { label: '6개월 초과 12개월 이하', min: 6, max: 12 },
    { label: '12개월 초과 24개월 이하', min: 12, max: 24 },
    { label: '24개월 초과', min: 24, max: Infinity },
  ];
}

function amountBand_(amount) {
  const hit = loanAmountBands_().find(band => amount > band.min && amount <= band.max);
  return hit ? hit.label : '';
}

function ltvBand_(ltv) {
  const hit = ltvBands_().find(band => ltv > band.min && ltv <= band.max);
  return hit ? hit.label : '';
}

function termBand_(term) {
  const hit = termBands_().find(band => term > band.min && term <= band.max);
  return hit ? hit.label : '';
}

function overdueBand_(row) {
  const y = Number(row['기준년']);
  const m = Number(row['기준월']);
  const repay = date_(row['상환일']);
  if (!y || !m || !repay) return '';
  const end = endOfMonth_(y, m);
  const days = Math.floor((end.getTime() - repay.getTime()) / 86400000);
  if (days <= 30) return '';
  const months = datedifMonths_(repay, end);
  return overdueTermBand_(months);
}

function overdueTermBand_(months) {
  if (months < 3) return '3개월 이하';
  if (months < 6) return '3개월 초과 6개월 이하';
  if (months < 12) return '6개월 초과 12개월 이하';
  if (months < 24) return '12개월 초과 24개월 이하';
  return '24개월 초과';
}

function samePeriod_(row, year, month) {
  return Number(row['기준년']) === Number(year) && Number(row['기준월']) === Number(month);
}

function matchBorrower_(row, borrower) {
  return text_(row['대출자구분']).trim() === borrower;
}

function matchProduct_(row, products) {
  const raw = text_(row['세부상품'] || row['상품분류_정리']);
  return products.includes(raw);
}

function normalizeInvestorType_(value) {
  const text = text_(value);
  if (text.includes('소득')) return '소득적격';
  if (text.includes('전문')) return '전문';
  if (text.includes('법인')) return '법인';
  if (text.includes('금융') || text.includes('여신')) return '금융기관';
  if (text.includes('온투')) return '온투업자';
  return text || '개인';
}

function normalizeProduct_(value) {
  const text = text_(value);
  if (text.includes('PF')) return '부동산 PF';
  if (text.includes('기타부동산')) return '기타부동산 담보대출';
  if (text.includes('근저당')) return '근저당질권';
  if (text.includes('부동산담보') || text.includes('부동산 담보')) return '부동산담보대출';
  if (text.includes('동산 담보')) return '동산 담보대출';
  if (text.includes('기타담보')) return '기타담보대출';
  if (text.includes('임차')) return '임차보증금 담보대출';
  if (text.includes('NPL')) return 'NPL채권 담보대출';
  if (text.includes('ABL')) return 'ABL 대출';
  if (text.includes('유가증권')) return '유가증권 담보대출';
  if (text.includes('매출채권') || text.includes('어음')) return '매출채권 담보대출';
  if (text.includes('개인사업자') && text.includes('신용')) return '개인사업자 신용';
  if (text.includes('법인') && text.includes('신용')) return '법인신용';
  if (text.includes('신용')) return '신용대출';
  return text || '기타';
}

function uniqueCount_(rows, key) {
  return new Set(rows.map(row => text_(row[key])).filter(Boolean)).size;
}

function uniqueCountNonzero_(rows, key) {
  return uniqueValues_(rows.map(row => row[key])).length;
}

function uniqueValues_(values) {
  return Array.from(new Set(values.map(text_).filter(validId_)));
}

function validId_(value) {
  const text = text_(value);
  const compact = text.replace(/[\s-]/g, '');
  return Boolean(compact) && !/^0+(\.0+)?$/.test(compact);
}

function sum_(rows, key) {
  return rows.reduce((total, row) => total + number_(row[key]), 0);
}

function groupSum_(rows, idKey, amountKey) {
  const out = {};
  rows.forEach(row => {
    const key = text_(row[idKey]);
    if (!validId_(key)) return;
    out[key] = (out[key] || 0) + number_(row[amountKey]);
  });
  return out;
}

function groupBy_(rows, keyFn) {
  const out = {};
  rows.forEach(row => {
    const key = keyFn(row);
    if (!out[key]) out[key] = [];
    out[key].push(row);
  });
  return out;
}

function sumEveryOther_(values, startIndex) {
  let total = 0;
  for (let i = startIndex; i < values.length; i += 2) total += number_(values[i]);
  return total;
}

function max_(values) {
  const nums = values.map(number_).filter(v => !isNaN(v));
  return nums.length ? Math.max(...nums) : 0;
}

function maxContractForValue_(rows, valueKey, targetValue) {
  const target = number_(targetValue);
  return max_(rows
    .filter(row => number_(row[valueKey]) === target)
    .map(row => number_(row['계약No'])));
}

function text_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function valueOr_(value, fallback) {
  return value === null || value === undefined || value === '' ? fallback : value;
}

function number_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const parsed = Number(String(value).replace(/,/g, ''));
  return isNaN(parsed) ? 0 : parsed;
}

function date_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400000));
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function period_(year, month) {
  return { year: Number(year), month: Number(month) };
}

function addPeriod_(period, deltaMonths) {
  const date = new Date(period.year, period.month - 1 + deltaMonths, 1);
  return { year: date.getFullYear(), month: date.getMonth() + 1 };
}

function endOfMonth_(year, month) {
  return new Date(year, month, 0);
}

function datedifMonths_(startDate, endDate) {
  let months = (endDate.getFullYear() - startDate.getFullYear()) * 12 + endDate.getMonth() - startDate.getMonth();
  if (endDate.getDate() < startDate.getDate()) months -= 1;
  return Math.max(months, 0);
}

function cleanProductName_(name) {
  return text_(name).replace('[', '').replace(']', '');
}

function maskName_(name) {
  const text = text_(name);
  if (text.length <= 1) return text;
  return text[0] + '*'.repeat(text.length - 1);
}

function businessNo_(id) {
  const text = text_(id).replace(/[^0-9]/g, '');
  return text.length >= 13 ? text.slice(0, 6) + '-' + text.slice(6, 13) : text;
}

function yyyymmdd_(value) {
  const parsed = date_(value);
  if (parsed) return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyyMMdd');
  return text_(value).replace(/[^0-9]/g, '').slice(0, 8);
}

function clone_(row) {
  const out = {};
  Object.keys(row).forEach(key => out[key] = row[key]);
  return out;
}
