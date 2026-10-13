const SHEET_NAME = '설문지 응답 시트1';
const STATUS = ['접수', '진행중', '검토', '완료'];   // 폼의 '진행 상태' 선택지와 똑같이

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('접수·진행 현황판')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function toYmd(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
}

function dayNum(ymd) {                       // 'yyyy-MM-dd' → 날짜 번호
  const p = String(ymd).split('-');
  if (p.length !== 3) return null;
  return Math.round(Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000);
}

function getBoard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('시트 탭 이름을 확인하세요: ' + SHEET_NAME);
  const tz = ss.getSpreadsheetTimeZone();
  const data = sheet.getDataRange().getValues();
  const today = dayNum(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));

  // 같은 접수번호가 여러 줄이면 '맨 아래 줄'이 현재 상태입니다.
  const latest = {};
  const count = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const no = String(r[2]).trim();
    if (!no) continue;
    latest[no] = r;
    count[no] = (count[no] || 0) + 1;
  }

  const items = Object.keys(latest).map(function (no) {
    const r = latest[no];
    const status = String(r[7]).trim();
    const due = toYmd(r[6], tz);
    const dn = dayNum(due);
    let dday = null;
    if (dn !== null) dday = dn - today;           // 음수면 납기 지남
    return { no: no, received: toYmd(r[1], tz), kind: String(r[3]), content: String(r[4]),
             manager: String(r[5]), due: due, status: status,
             dday: dday, updates: count[no], note: String(r[8] || '') };
  });
  items.sort(function (a, b) { return (a.due || '9').localeCompare(b.due || '9'); });
  return { statuses: STATUS, items: items };
}
