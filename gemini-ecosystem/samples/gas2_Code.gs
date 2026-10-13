const SHEET_NAME = '설문지 응답 시트1';

// 품목별 단위와 안전재고(이 수량보다 적으면 부족 표시). 내 회사 품목 이름으로 바꿔 쓰세요.
const ITEMS = {
  '주원료 A': { unit: 'kg', safe: 200 },
  '주원료 B': { unit: 'kg', safe: 100 },
  '첨가제': { unit: 'kg', safe: 50 },
  '포장 포대': { unit: '개', safe: 300 },
  '라벨': { unit: '장', safe: 1000 }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('원부자재 재고 현황')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('시트 탭 이름을 확인하세요: ' + SHEET_NAME);
  const tz = ss.getSpreadsheetTimeZone();
  const data = sheet.getDataRange().getValues();

  const sum = {};
  Object.keys(ITEMS).forEach(function (k) { sum[k] = { inQty: 0, outQty: 0 }; });
  const logs = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const kind = String(r[2]).trim();
    const name = String(r[3]).trim();
    const qty = Number(r[4]);
    if (!name || !(qty > 0)) continue;
    if (kind !== '입고' && kind !== '출고') continue;
    if (!sum[name]) sum[name] = { inQty: 0, outQty: 0 };
    if (kind === '입고') sum[name].inQty += qty; else sum[name].outQty += qty;
    logs.push({
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
      kind: kind, name: name, qty: qty,
      manager: String(r[5]), note: String(r[6])
    });
  }

  const items = Object.keys(sum).map(function (name) {
    const meta = ITEMS[name] || { unit: '', safe: 0 };
    const now = sum[name].inQty - sum[name].outQty;
    let status = '정상';
    if (meta.safe > 0 && now < meta.safe) status = now < meta.safe / 2 ? '긴급' : '부족';
    return { name: name, unit: meta.unit, safe: meta.safe, now: now,
             inQty: sum[name].inQty, outQty: sum[name].outQty, status: status };
  });
  logs.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return { items: items, logs: logs.slice(0, 30) };
}
