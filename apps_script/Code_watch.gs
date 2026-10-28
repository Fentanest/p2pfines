/**
 * Jjukfund online watch-report calculator.
 *
 * This file is installed only in the 상시감시보고서 spreadsheet.
 * Trace path: raw tabs -> 계산_* tabs -> (상시) report tabs -> 검증/실행로그
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

const WATCH_REPORTS = ['(상시)1-1', '(상시)1-3', '(상시)2-2', '(상시)2-3', '(상시)3-1', '(상시)3-2'];
const WATCH_PERIOD_REPORTS = ['(상시)1-1', '(상시)2-2', '(상시)3-1', '(상시)3-2'];
const RUN_SHEET = '실행';
const RUN_CHECKBOX_CELL = 'B7';
const WATCH_PRODUCTS = ['부동산 PF', '부동산담보대출', '동산 담보대출', '매출채권 담보대출', '신용대출', '개인사업자 신용', '법인신용'];
const WATCH_AGES = ['19세 미만', '19세 이상 30세 미만', '30세 이상 40세 미만', '40세 이상 50세 미만', '50세 이상 60세 미만', '60세 이상 70세 미만', '70세 이상 80세 미만', '80세 이상 90세 미만', '90세 이상'];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui
    .createMenu('Jjukfund')
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
  ss.toast('보고서 재계산을 시작했습니다.', 'Jjukfund', 5);
  try {
    recalculateAllReports();
    setupRunButton({ silent: true });
    const seconds = Math.round((new Date().getTime() - startedAt.getTime()) / 1000);
    ss.toast('보고서 재계산을 완료했습니다. (' + seconds + '초)', 'Jjukfund', 10);
  } catch (error) {
    ss.toast('보고서 재계산 실패: ' + error.message, 'Jjukfund', 10);
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
    .setValue('Jjukfund 보고서 실행')
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
    ['상단 메뉴', 'Jjukfund > 보고서 재계산', '시트를 다시 열면 메뉴 표시'],
  ]);
  sheet.getRange('B9:D9')
    .setFontWeight('bold')
    .setBackground('#d7dce3');
  sheet.getRange('B9:D11')
    .setBorder(true, true, true, true, true, true, '#9aa8b8', SpreadsheetApp.BorderStyle.SOLID);

  if (!options || !options.silent) {
    ss.setActiveSheet(sheet);
    ss.toast('실행 버튼을 준비했습니다.', 'Jjukfund', 5);
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
  const period = readPeriod_(ss, settings, WATCH_REPORTS);
  const year = period.year;
  const month = period.month;

  const rawContractAll = readTable_(ss, RAW.CONTRACT_ALL);
  const rawInvestor = readTable_(ss, RAW.INVESTOR);
  const rawContractSnapshot = readTable_(ss, RAW.CONTRACT_SNAPSHOT);

  writeReportPeriod_(ss, year, month);
  writeReadmePeriod_(ss, year, month);
  const normalized = normalizeSources_(ss, rawContractAll, rawInvestor, rawContractSnapshot, year, month);
  clearReportOutputRanges_(ss);

  recalcWatch11_(ss, normalized.contractAll, normalized.investor, year, month);
  recalcWatch13_(ss, normalized.contractSnapshot, year, month);
  recalcWatch31_(ss, normalized.investor, year, month);
  recalcWatch22_(ss, normalized.investor, year, month);
  recalcWatch23_(ss, normalized.contractSnapshot, year, month);
  recalcWatch32_(ss, normalized.contractSnapshot, year, month);

  writeVerification_(ss, normalized, year, month);
  appendLog_(ss, startedAt, new Date(), year, month, normalized);
}

function writeReportPeriod_(ss, year, month) {
  WATCH_PERIOD_REPORTS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.getRange(1, 1, 1, 2).setValues([[year, month]]);
  });
}

function clearReportOutputRanges_(ss) {
  watchOutputRanges_().forEach(spec => clearOutputRange_(ss, spec));
}

function watchOutputRanges_() {
  return [
    { sheetName: '(상시)1-1', row: 6, col: 1, rows: 'max', cols: 25 },
    { sheetName: '(상시)1-3', row: 6, col: 1, rows: 1, cols: 49 },
    { sheetName: '(상시)1-3', row: 15, col: 1, rows: 'max', cols: 4 },
    { sheetName: '(상시)2-2', row: 8, col: 1, rows: 30, cols: 20 },
    { sheetName: '(상시)2-2', row: 40, col: 1, rows: 'max', cols: 20 },
    { sheetName: '(상시)2-3', row: 3, col: 1, rows: 99, cols: 72 },
    { sheetName: '(상시)2-3', row: 110, col: 1, rows: 'max', cols: 8 },
    { sheetName: '(상시)3-1', row: 5, col: 1, rows: 1, cols: 19 },
    { sheetName: '(상시)3-1', row: 8, col: 1, rows: 'max', cols: 19 },
    { sheetName: '(상시)3-2', row: 10, col: 1, rows: 'max', cols: 16 },
  ];
}

function normalizeSources_(ss, contractAll, investor, contractSnapshot, year, month) {
  const contractAllNorm = normalizeContractAll_(contractAll, year, month);
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

function normalizeContractAll_(rows, year, month) {
  const baseSerial = endOfMonthSerial_(year, month);
  const recentCompleteStartSerial = addMonthsSerial_(baseSerial, -6) + 1;
  const excluded = new Set(['385', '460', '471', '478', '562', '814', '817']);
  return rows.map(row => {
    const out = clone_(row);
    applyAdjustedStatus_(out);
    const contractDate = dateSerial_(out['계약일']);
    const completeDate = dateSerial_(out['완납일']);
    const no = text_(out['계약No']);
    const inferredActive = contractDate &&
      contractDate <= baseSerial &&
      !excluded.has(no) &&
      (!completeDate || completeDate > baseSerial || (completeDate >= recentCompleteStartSerial && completeDate <= baseSerial));
    const explicitActive = explicitBoolean_(out['(상시)1-1']);
    const active = explicitActive === null ? inferredActive : explicitActive;
    out['(상시)1-1대상'] = active ? 'TRUE' : 'FALSE';
    out['상품명_정리'] = cleanProductName_(out['상품명']);
    return out;
  });
}

function normalizeInvestor_(rows, contractByNo, year, month) {
  return rows.map(row => {
    const out = clone_(row);
    out['기준년'] = yearNumber_(out['기준년'], year);
    out['기준월'] = monthNumber_(out['기준월'], month);
    const contract = contractByNo[text_(out['계약번호'])] || {};
    out['대출자구분'] = valueOr_(out['대출자구분'], contract['대출자구분']);
    out['(상시)3-1'] = [out['기준년'], out['기준월'], out['주민등록번호']].map(text_).join('-');
    out['투자자구분_정리'] = normalizeInvestorType_(out['투자자구분']);
    out['상품분류_정리'] = normalizeProduct_(out['세부상품'] || contract['세부상품']);
    out['투자잔액_숫자'] = number_(out['투자잔액']);
    return out;
  });
}

function normalizeContractSnapshot_(rows, contractByNo, year, month) {
  return rows.map(row => {
    const out = clone_(row);
    const contract = contractByNo[text_(out['계약No'])] || {};
    applyAdjustedStatus_(out, contract);
    out['기준년'] = yearNumber_(out['기준년'], year);
    out['기준월'] = monthNumber_(out['기준월'], month);
    out['잔액_숫자'] = number_(out['잔액']);
    out['총대출액_숫자'] = number_(out['총대출액']);
    out['상품분류_정리'] = normalizeProduct_(out['세부상품']);
    out['월_계약No'] = [out['기준년'], out['기준월'], out['계약No']].map(text_).join('-');
    out['월_계약No_상태'] = [out['기준년'], out['기준월'], out['계약No'], out['상태']].map(text_).join('-');
    out['월_주민등록'] = [out['기준년'], out['기준월'], out['주민등록']].map(text_).join('-');
    return out;
  });
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

function writeVerification_(ss, normalized, year, month) {
  const currentInvestor = normalized.investor.filter(row => samePeriod_(row, year, month));
  const currentSnapshot = normalized.contractSnapshot.filter(row => samePeriod_(row, year, month));
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
    ['기준일계약 보유 기간', snapshotPeriods.join(', '), '상시감시보고서 월별 추적 원천'],
    ['당월 투자잔액 합계', sum_(currentInvestor, '투자잔액_숫자'), '(상시)2-2/(상시)3-1 원천'],
    ['당월 대출잔액 합계', sum_(currentSnapshot, '잔액_숫자'), '(상시)1-3/(상시)2-3/(상시)3-2 원천'],
    ['연체 대출잔액 합계', sum_(currentSnapshot.filter(row => text_(row['상태']) === '연체'), '잔액_숫자'), '(상시)1-3 원천'],
  ];
  writeValues_(ss, CALC.VERIFY, rows);
}

function recalcWatch11_(ss, contractAll, investor, year, month) {
  const currentSnapshots = readTable_(ss, CALC.CONTRACT_SNAPSHOT).filter(row => samePeriod_(row, year, month));
  const snapshotByNo = {};
  currentSnapshots.forEach(row => {
    snapshotByNo[text_(row['계약No'])] = row;
  });
  const investorByPeriodProduct = groupBy_(investor, row => {
    return [yearNumber_(row['기준년'], year), monthNumber_(row['기준월'], month), text_(row['상품명'])].join('|');
  });
  const rows = contractAll
    .filter(row => row['(상시)1-1대상'] === 'TRUE')
    .map((row, idx) => {
      const product = cleanProductName_(row['상품명']);
      const no = text_(row['계약No']);
      const startSerial = dateSerial_(row['계약일']);
      const startParts = serialYearMonth_(startSerial, year, month);
      const investors = investorByPeriodProduct[[startParts.year, startParts.month, product].join('|')] || [];
      const snapshot = snapshotByNo[no] || {};
      const limit = number_(row['한도'] || row['총대출']);
      const feeType = text_(row['수수료구분']);
      const paidAmount = feeType === '후취' ? limit : limit - number_(row['취급수수료']) * 1.1;
      const borrowerType = row['대출자구분'] || snapshot['대출자구분'] || '';
      const borrowerName = row['성명'] || snapshot['성명'] || '';
      const borrowerId = row['주민등록'] || snapshot['주민등록'] || '';
      const investorCounts = ['개인', '소득적격', '전문', '법인', '금융기관']
        .map(type => investorTypeCount_(investors, type));
      const expectedInvestorCount = number_(row['투자자수']);
      const investorTotal = expectedInvestorCount && expectedInvestorCount === investorCounts.reduce((a, b) => a + b, 0)
        ? expectedInvestorCount
        : '';
      return [
        idx + 1,
        product,
        product.replace('줌', '').replace('호', ''),
        detailProductLabel_(row['세부상품'] || row['상품'] || row['상품명']),
        text_(row['유형'] || row['최근대출']).includes('리파이낸싱') ? '여' : '부',
        watch11StatusLabel_(idx),
        no,
        limit,
        paidAmount || '',
        limit,
        snapshot['잔액_숫자'] === undefined ? '' : number_(snapshot['잔액_숫자']),
        '',
        row['이율'] || '',
        formatDate_(row['계약일']),
        formatDate_(row['만기일']),
        formatDate_(row['완납일']),
        borrowerType,
        borrowerType === '개인' ? maskName_(borrowerName) : borrowerName,
        identifier_(borrowerId, borrowerType),
        ...investorCounts,
        investorTotal,
      ];
    });
  const total = Array(25).fill('');
  total[0] = '합계';
  total[4] = rows.filter(row => row[4] === '여').length;
  total[5] = rows.filter(row => row[5] === '정상상환중' || row[5] === '연체중').length;
  [8, 9, 10, 11, 19, 20, 21, 22, 23, 24].forEach(idx => total[idx] = rows.reduce((sum, row) => sum + number_(row[idx]), 0));
  total[12] = 0;
  total[13] = '하이픈("-") 포함\nex. 2024-01-01';
  total[14] = '하이픈("-") 포함\nex. 2024-03-31';
  total[15] = '하이픈("-") 포함\nex. 2024-03-31';

  const sheet = ensureSheet_(ss, '(상시)1-1');
  sheet.getRange(5, 20, 1, 6).setValues([['일반', '소득적격', '전문', '일반', '여신금융기관', '']]);
  sheet.getRange(6, 1, Math.max(sheet.getMaxRows() - 5, 1), 25).clearContent();
  sheet.getRange(7, 14, Math.max(sheet.getMaxRows() - 6, 1), 3).setNumberFormat('@');
  writeMatrix_(ss, '(상시)1-1', 6, 1, [total]);
  if (rows.length) writeMatrix_(ss, '(상시)1-1', 7, 1, rows);
}

function recalcWatch13_(ss, snapshots, year, month) {
  const contractAll = readTable_(ss, CALC.CONTRACT_ALL);
  const row = [year];
  for (let m = 1; m <= 12; m++) {
    if (m > month) {
      row.push('', '', '', '');
      continue;
    }
    const monthEnd = endOfMonthSerial_(year, m);
    const cumulative = contractAll
      .filter(item => {
        const serial = dateSerial_(item['계약일']);
        return serial && serial <= monthEnd;
      })
      .reduce((sum, item) => sum + number_(item['총대출'] || item['한도']), 0);
    const rowsForMonth = snapshots.filter(row => Number(row['기준년']) === year && Number(row['기준월']) === m);
    if (!rowsForMonth.length) {
      row.push(cumulative || '', '', '', '');
      continue;
    }
    const balance = sum_(rowsForMonth, '잔액_숫자');
    const overdue = sum_(rowsForMonth.filter(row => text_(row['상태']).includes('연체')), '잔액_숫자');
    row.push(cumulative, balance, overdue, balance ? overdue / balance * 100 : 0);
  }
  writeMatrix_(ss, '(상시)1-3', 6, 1, [row]);
}

function recalcWatch22_(ss, investor, year, month) {
  const source = buildWatch31Rows_(investor, year, month).objects;
  const rows = [];
  const groupSpecs = [
    { label: '개인일반', predicate: row => row['투자자구분 표시'] === '개인일반' },
    { label: '소득적격', predicate: row => row['투자자구분 표시'] === '소득적격' },
    { label: '개인 합계', predicate: row => row['투자자구분 표시'] === '개인일반' || row['투자자구분 표시'] === '소득적격' },
  ];
  groupSpecs.forEach(group => {
    WATCH_AGES.forEach((age, ageIdx) => {
      const filtered = source.filter(row => group.predicate(row) && row['연령'] === age);
      const values = WATCH_PRODUCTS.flatMap(product => [sum_(filtered, product), countPositive_(filtered, product)]);
      const totalAmount = values.filter((_, idx) => idx % 2 === 0).reduce((sum, value) => sum + number_(value), 0);
      const totalCount = values.filter((_, idx) => idx % 2 === 1).reduce((sum, value) => sum + number_(value), 0);
      rows.push([ageIdx === 0 ? group.label : '', age, ...values, totalAmount, totalCount, sum_(filtered, '연체잔액'), countPositive_(filtered, '연체잔액')]);
    });
    const block = rows.slice(rows.length - WATCH_AGES.length);
    const subtotal = ['', group.label === '개인 합계' ? '합계' : '소계'];
    for (let idx = 2; idx < 20; idx++) {
      subtotal[idx] = block.reduce((sum, row) => sum + number_(row[idx]), 0);
    }
    rows.push(subtotal);
  });
  writeMatrix_(ss, '(상시)2-2', 8, 1, rows);
}

function recalcWatch23_(ss, snapshots, year, month) {
  const contractAll = readTable_(ss, CALC.CONTRACT_ALL);
  const contractByNo = {};
  contractAll.forEach(row => contractByNo[text_(row['계약No'])] = row);
  const matrix = Array.from({ length: 99 }, () => Array(72).fill(''));
  for (let m = 1; m <= 12; m++) {
    const offset = (m - 1) * 6;
    matrix[0][offset] = year;
    matrix[0][offset + 1] = m;
    ['계약번호', '대출잔액', '연체잔액', '연체율', '상품수', '투자자수'].forEach((header, idx) => matrix[1][offset + idx] = header);
    matrix[2][offset] = '합계';
    if (m > month) continue;
    const rows = snapshots.filter(row => samePeriod_(row, year, m) && text_(row['리파이낸싱 여부']).includes('리파이낸싱'));
    rows.forEach((row, idx) => {
      if (idx >= 96) return;
      const line = 3 + idx;
      const balance = number_(row['잔액_숫자']);
      const overdue = isOverdueStatus_(row['상태']) ? balance : '';
      const contract = contractByNo[text_(row['계약No'])] || {};
      matrix[line][offset] = row['계약No'];
      matrix[line][offset + 1] = balance;
      matrix[line][offset + 2] = overdue;
      matrix[line][offset + 3] = overdue && balance ? number_(overdue) / balance * 100 : '';
      matrix[line][offset + 4] = '';
      matrix[line][offset + 5] = contract['투자자수'] || '';
    });
    matrix[2][offset + 1] = rows.reduce((sum, row) => sum + number_(row['잔액_숫자']), 0);
    matrix[2][offset + 2] = rows.reduce((sum, row) => sum + (isOverdueStatus_(row['상태']) ? number_(row['잔액_숫자']) : 0), 0);
    matrix[2][offset + 3] = matrix[2][offset + 1] ? matrix[2][offset + 2] / matrix[2][offset + 1] * 100 : 0;
    matrix[2][offset + 4] = rows.length;
    matrix[2][offset + 5] = rows.reduce((sum, row) => sum + number_((contractByNo[text_(row['계약No'])] || {})['투자자수']), 0);
  }
  writeMatrix_(ss, '(상시)2-3', 3, 1, matrix);
}

function recalcWatch31_(ss, investor, year, month) {
  const built = buildWatch31Rows_(investor, year, month);
  const sheet = ensureSheet_(ss, '(상시)3-1');
  sheet.getRange(7, 1, 1, 19).setValues([[
    '주민등록번호', '성명', '나이', '연령', '연체잔액', '투자자구분', '투자자구분', '성명',
    '투자자 식별정보', '관리번호', '부동산 PF', '부동산담보대출', '동산 담보대출',
    '매출채권 담보대출', '신용대출', '개인사업자\n신용', '법인신용', '합계', '상품건수'
  ]]);
  const summary = Array(19).fill('');
  summary[0] = built.objects.length;
  WATCH_PRODUCTS.forEach((product, idx) => summary[10 + idx] = sum_(built.objects, product));
  summary[17] = WATCH_PRODUCTS.reduce((total, product) => total + sum_(built.objects, product), 0);
  summary[18] = sum_(built.objects, '상품건수');
  writeMatrix_(ss, '(상시)3-1', 5, 1, [summary]);
  if (built.rows.length) writeMatrix_(ss, '(상시)3-1', 8, 1, built.rows);
}

function buildWatch31Rows_(investor, year, month) {
  const firstByInvestor = {};
  investor.forEach(row => {
    const id = text_(row['주민등록번호']);
    if (id && !firstByInvestor[id]) firstByInvestor[id] = row;
  });
  const filtered = investor.filter(row => samePeriod_(row, year, month));
  const sameMonth = investor.filter(row => monthNumber_(row['기준월'], month) === month);
  const sameMonthProductSums = {};
  sameMonth.forEach(row => {
    const id = text_(row['주민등록번호']);
    const product = row['상품분류_정리'];
    if (!id || !product) return;
    const key = id + '|' + product;
    sameMonthProductSums[key] = (sameMonthProductSums[key] || 0) + number_(row['투자잔액_숫자']);
  });
  const byInvestor = {};
  filtered.forEach(row => {
    const id = text_(row['주민등록번호']);
    if (!id) return;
    const first = firstByInvestor[id] || row;
    if (!byInvestor[id]) {
      byInvestor[id] = { id, name: first['이름'] || '', age: first['나이'] || '', type: row['투자자구분_정리'], memberNo: row['회원번호'] || '', overdue: 0, products: {}, count: 0 };
    }
    byInvestor[id].count += 1;
    const amount = number_(row['투자잔액_숫자']);
    const product = row['상품분류_정리'];
    byInvestor[id].products[product] = (byInvestor[id].products[product] || 0) + amount;
    if (isOverdueStatus_(row['연체여부'])) byInvestor[id].overdue += amount;
  });
  Object.values(byInvestor).forEach(item => {
    Object.keys(item.products).forEach(product => {
      item.products[product] = sameMonthProductSums[item.id + '|' + product] || item.products[product];
    });
  });
  const objects = Object.values(byInvestor).map(item => {
    const productValues = WATCH_PRODUCTS.map(product => item.products[product] || '');
    const out = {
      '주민등록번호': item.id,
      '성명': item.name,
      '나이': item.age,
      '연령': ageBucket_(item.age),
      '연체잔액': item.overdue || '',
      '투자자구분': item.type,
      '투자자구분 표시': investorTypeLabel_(item.type),
      '표시명': item.type === '법인' ? item.name : maskName_(item.name),
      '투자자 식별정보': identifier_(item.id, item.type),
      '관리번호': item.memberNo,
      '합계': productValues.reduce((a, b) => a + number_(b), 0),
      '상품건수': item.count,
    };
    WATCH_PRODUCTS.forEach((product, idx) => out[product] = productValues[idx]);
    return out;
  });
  const headers = ['주민등록번호', '성명', '나이', '연령', '연체잔액', '투자자구분', '투자자구분 표시', '표시명', '투자자 식별정보', '관리번호', ...WATCH_PRODUCTS, '합계', '상품건수'];
  return { headers, objects, rows: objects.map(row => headers.map(header => row[header])) };
}

function recalcWatch32_(ss, snapshots, year, month) {
  const filtered = snapshots.filter(row => samePeriod_(row, year, month) && number_(row['잔액_숫자']) !== 0);
  const byBorrower = {};
  filtered.forEach(row => {
    const id = text_(row['주민등록']);
    if (!id) return;
    if (!byBorrower[id]) {
      byBorrower[id] = { id, type: row['대출자구분'] || '', name: row['성명'] || '', products: {}, count: 0 };
    }
    byBorrower[id].count += 1;
    const product = row['상품분류_정리'];
    byBorrower[id].products[product] = (byBorrower[id].products[product] || 0) + number_(row['잔액_숫자']);
  });
  const rows = Object.values(byBorrower).map((item, idx) => {
    const productValues = WATCH_PRODUCTS.map(product => item.products[product] || '');
    return [idx + 1, item.type, item.type === '개인' ? maskName_(item.name) : item.name, identifier_(item.id, item.type), '', ...productValues, productValues.reduce((a, b) => a + number_(b), 0), item.count];
  });
  const summary = Array(14).fill('');
  summary[0] = '합계';
  WATCH_PRODUCTS.forEach((product, idx) => summary[5 + idx] = rows.reduce((sum, row) => sum + number_(row[5 + idx]), 0));
  summary[12] = rows.reduce((sum, row) => sum + number_(row[12]), 0);
  summary[13] = rows.reduce((sum, row) => sum + number_(row[13]), 0);
  const sheet = ensureSheet_(ss, '(상시)3-2');
  sheet.getRange(9, 1, 1, 14).setValues([[
    '구분', '차입자 구분', '성명', '차입자 식별정보', '고유관리번호', '부동산 PF',
    '부동산담보', '기타\n(어음·매출채권 제외)', '어음·매출채권 담보',
    '개인 신용\n(개인사업자 신용 제외)', '개인사업자\n신용', '법인신용', '합계', '상품건수'
  ]]);
  writeMatrix_(ss, '(상시)3-2', 10, 1, [summary]);
  if (rows.length) writeMatrix_(ss, '(상시)3-2', 11, 1, rows);
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

function writeValues_(ss, sheetName, rows) {
  const sheet = ensureSheet_(ss, sheetName);
  sheet.clearContents();
  if (rows.length) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}

function writeMatrix_(ss, sheetName, startRow, startCol, rows) {
  if (!rows.length) return;
  const sheet = ensureSheet_(ss, sheetName);
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const normalized = rows.map(row => {
    const out = row.slice();
    while (out.length < width) out.push('');
    return out;
  });
  sheet.getRange(startRow, startCol, normalized.length, width).setValues(normalized);
}

function ensureSheet_(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function samePeriod_(row, year, month) {
  return yearNumber_(row['기준년'], 0) === Number(year) && monthNumber_(row['기준월'], 0) === Number(month);
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
  if (text.includes('부동산담보') || text.includes('부동산 담보')) return '부동산담보대출';
  if (text.includes('매출채권') || text.includes('어음')) return '매출채권 담보대출';
  if (text.includes('개인사업자') && text.includes('신용')) return '개인사업자 신용';
  if (text.includes('법인') && text.includes('신용')) return '법인신용';
  if (text.includes('신용')) return '신용대출';
  if (text.includes('동산') || text.includes('담보')) return '동산 담보대출';
  return text || '동산 담보대출';
}

function investorTypeLabel_(type) {
  const map = { '개인': '개인일반', '법인': '법인일반', '소득적격': '소득적격', '전문': '전문투자자', '금융기관': '여신금융기관', '온투업자': '온라인투자연계금융업자' };
  return map[type] || type;
}

function investorTypeCount_(rows, type) {
  return rows.filter(row => row['투자자구분_정리'] === type).length;
}

function uniqueCount_(rows, key) {
  return new Set(rows.map(row => text_(row[key])).filter(Boolean)).size;
}

function countPositive_(rows, key) {
  return rows.filter(row => number_(row[key]) > 0).length;
}

function explicitBoolean_(value) {
  if (value === true) return true;
  if (value === false) return false;
  const text = text_(value).toLowerCase();
  if (['true', 'y', 'yes', '1', '대상', '여'].includes(text)) return true;
  if (['false', 'n', 'no', '0', '제외', '부'].includes(text)) return false;
  return null;
}

function isOverdueStatus_(value) {
  return text_(value).replace(/[()]/g, '').trim() === '연체';
}

function sum_(rows, key) {
  return rows.reduce((total, row) => total + number_(row[key]), 0);
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

function yearNumber_(value, fallback) {
  if (value === null || value === undefined || value === '') return Number(fallback || 0);
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const year = value.getFullYear();
    return year > 1909 ? year : Number(fallback || 0);
  }
  const parsed = Number(String(value).replace(/,/g, '').trim());
  return isNaN(parsed) ? Number(fallback || 0) : Math.round(parsed);
}

function monthNumber_(value, fallback) {
  if (value === null || value === undefined || value === '') return Number(fallback || 0);
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const year = value.getFullYear();
    if (year > 1909) return value.getMonth() + 1;
    return Math.round(dateSerial_(value));
  }
  const parsed = Number(String(value).replace(/,/g, '').trim());
  return isNaN(parsed) ? Number(fallback || 0) : Math.round(parsed);
}

function dateSerial_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return Math.floor(value);
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Math.round(value.getTime() / 86400000 + 25569);
  }
  const text = String(value).trim();
  if (!text) return null;
  const numeric = Number(text.replace(/,/g, ''));
  if (!isNaN(numeric)) return Math.floor(numeric);
  const match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) return dateToSerial_(Number(match[1]), Number(match[2]), Number(match[3]));
  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return null;
  return Math.round(parsed.getTime() / 86400000 + 25569);
}

function dateToSerial_(year, month, day) {
  return Math.round(Date.UTC(year, month - 1, day) / 86400000 + 25569);
}

function endOfMonthSerial_(year, month) {
  return Math.round(Date.UTC(year, month, 0) / 86400000 + 25569);
}

function addMonthsSerial_(serial, delta) {
  const date = serialToDate_(serial);
  const targetYear = date.getUTCFullYear();
  const targetMonthIndex = date.getUTCMonth() + delta;
  const lastDay = new Date(Date.UTC(targetYear, targetMonthIndex + 1, 0)).getUTCDate();
  const day = Math.min(date.getUTCDate(), lastDay);
  const shifted = new Date(Date.UTC(targetYear, targetMonthIndex, day));
  return Math.round(shifted.getTime() / 86400000 + 25569);
}

function serialToDate_(serial) {
  return new Date(Math.round((Number(serial) - 25569) * 86400000));
}

function serialYearMonth_(serial, fallbackYear, fallbackMonth) {
  if (!serial) return { year: fallbackYear, month: fallbackMonth };
  const date = serialToDate_(serial);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1 };
}

function formatDate_(value) {
  const serial = dateSerial_(value);
  if (!serial) return '';
  const date = serialToDate_(serial);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function date_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400000));
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function endOfMonth_(year, month) {
  return new Date(year, month, 0);
}

function addMonths_(date, delta) {
  return new Date(date.getFullYear(), date.getMonth() + delta, date.getDate());
}

function cleanProductName_(name) {
  return text_(name).replace('[', '').replace(']', '');
}

function detailProductLabel_(name) {
  return text_(name).replace(/\s*대출$/, '');
}

function reportLoanStatus_(row) {
  const status = text_(row['상태']);
  if (status.includes('매각')) return '채권매각';
  if (dateSerial_(row['완납일'])) return '완제';
  if (status.includes('연체')) return '연체중';
  if (status.includes('정상')) return '정상상환중';
  return status;
}

function watch11StatusLabel_(idx) {
  return ['정상상환중', '완제', '연체중', '채권매각', '기타'][idx] || '';
}

function maskName_(name) {
  const text = text_(name);
  if (text.length <= 1) return text;
  return text[0] + '*'.repeat(text.length - 1);
}

function identifier_(id, type) {
  const text = text_(id);
  if (type === '법인') return businessNo_(text);
  const digits = text.replace(/[^0-9]/g, '');
  return digits.length >= 7 ? digits.slice(0, 6) + '-' + digits.slice(6, 7) : text;
}

function businessNo_(id) {
  const text = text_(id).replace(/[^0-9]/g, '');
  return text.length >= 13 ? text.slice(0, 6) + '-' + text.slice(6, 13) : text;
}

function ageBucket_(age) {
  if (age === null || age === undefined || age === '') return '19세 미만';
  const n = number_(age);
  if (n < 19) return '19세 미만';
  if (n < 30) return '19세 이상 30세 미만';
  if (n < 40) return '30세 이상 40세 미만';
  if (n < 50) return '40세 이상 50세 미만';
  if (n < 60) return '50세 이상 60세 미만';
  if (n < 70) return '60세 이상 70세 미만';
  if (n < 80) return '70세 이상 80세 미만';
  if (n < 90) return '80세 이상 90세 미만';
  return '90세 이상';
}

function clone_(row) {
  const out = {};
  Object.keys(row).forEach(key => out[key] = row[key]);
  return out;
}
