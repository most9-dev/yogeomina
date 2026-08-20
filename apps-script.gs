// ============================================================
// 라이브 주문서 — 구글 시트 연동 스크립트
// 사용법: 구글 시트에서 [확장 프로그램 → Apps Script]에 이 코드를
// 통째로 붙여넣고, setup 함수를 한 번 실행한 뒤 웹 앱으로 배포하세요.
// (자세한 순서는 설치가이드.md 참고)
// ============================================================

const SHEET_PRODUCTS = '상품목록';
const SHEET_ORDERS = '주문접수';

// 최초 1회 실행: 필요한 탭과 머리글을 자동으로 만들어 줍니다.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let p = ss.getSheetByName(SHEET_PRODUCTS);
  if (!p) {
    p = ss.insertSheet(SHEET_PRODUCTS);
    p.appendRow(['상품번호', '상품명', '가격', '칼라(쉼표로 구분)', '사이즈(쉼표로 구분)', '재고(총수량)']);
    p.appendRow(['101', '니트 가디건', 29000, '아이보리,블랙,핑크', 'Free', 10]);
    p.appendRow(['102', '와이드 팬츠', 24000, '베이지,차콜', 'S,M,L', '']);
    p.setFrozenRows(1);
  }

  let o = ss.getSheetByName(SHEET_ORDERS);
  if (!o) {
    o = ss.insertSheet(SHEET_ORDERS);
    o.appendRow(['주문시각', '유튜브닉네임', '수령인', '연락처', '주소', '배송메모', '배송지역', '입금자명',
                 '상품번호', '상품명', '칼라', '사이즈', '수량', '금액',
                 '입금할 총액', '입금확인']);
    o.setFrozenRows(1);
  }
}

// [상품목록]에 판매수량/남은재고 열을 자동으로 만들어 줍니다.
// 실행하면 G열·H열에 수식이 채워지고, 이후 주문이 들어올 때마다 자동 갱신됩니다.
// 상품 줄을 새로 추가한 뒤에는 이 함수를 한 번 더 실행해 주세요.
function addStockColumns() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const last = sh.getLastRow();
  sh.getRange('G1').setValue('판매수량');
  sh.getRange('H1').setValue('남은재고');
  for (let r = 2; r <= last; r++) {
    if (sh.getRange('A' + r).getValue() === '') continue;
    sh.getRange('G' + r).setFormula(
      '=SUMPRODUCT((주문접수!$I$2:$I$1000&""=$A' + r + '&"")*주문접수!$M$2:$M$1000)');
    sh.getRange('H' + r).setFormula(
      '=IF($F' + r + '="","무제한",IFERROR($F' + r + '-$G' + r + ',"옵션별"))');
  }
}

// 주문서 페이지가 열릴 때 상품 목록을 내려줍니다.
// ?action=lookup 으로 호출되면 주문 조회로 동작합니다.
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'lookup') {
    return lookupOrders(e.parameter.ytname, e.parameter.phone);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  const sold = soldByProduct(ss);
  const products = {};
  for (let i = 1; i < rows.length; i++) {
    const [no, name, price, colors, sizes, stock] = rows[i];
    if (!no || !name) continue;
    const key = String(no).trim();
    const p = {
      name: String(name).trim(),
      price: Number(price) || 0,
      colors: String(colors || '').split(',').map(s => s.trim()).filter(Boolean),
      sizes: String(sizes || '').split(',').map(s => s.trim()).filter(Boolean),
    };
    // 재고 계산: 빈칸=무제한, 숫자 하나=상품 전체, 쉼표 목록=옵션별(칼라 또는 사이즈 순서)
    const st = parseStock(stock, p.colors, p.sizes);
    if (st) {
      if (st.by === 'total') {
        p.stock = Math.max(0, st.total - (sold.total[key] || 0));
      } else {
        const opts = st.by === 'color' ? p.colors : p.sizes;
        const soldMap = st.by === 'color' ? sold.byColor : sold.bySize;
        p.stockBy = st.by;
        p.stockList = opts.map(function (o, idx) {
          return Math.max(0, st.list[idx] - (soldMap[key + '|' + o] || 0));
        });
        p.stock = p.stockList.reduce(function (a, b) { return a + b; }, 0);
      }
    }
    products[key] = p;
  }
  return ContentService.createTextOutput(JSON.stringify(products))
    .setMimeType(ContentService.MimeType.JSON);
}

// 재고 칸 해석: '' → null(무제한) / '10' → 전체 10개 / '3,5,2' → 옵션 순서별
function parseStock(stock, colors, sizes) {
  const s = String(stock === null || stock === undefined ? '' : stock).trim();
  if (s === '') return null;
  if (s.indexOf(',') === -1) {
    const n = Number(s);
    return isNaN(n) ? null : { by: 'total', total: n };
  }
  const list = s.split(',').map(function (x) { return Number(String(x).trim()); });
  if (list.some(function (n) { return isNaN(n); })) return null;
  if (colors.length === list.length) return { by: 'color', list: list };
  if (sizes.length === list.length) return { by: 'size', list: list };
  // 옵션 개수와 안 맞으면 합계를 전체 재고로 취급
  return { by: 'total', total: list.reduce(function (a, b) { return a + b; }, 0) };
}

// [주문접수]에 쌓인 주문에서 상품별/옵션별 판매 수량을 합산합니다.
function soldByProduct(ss) {
  const sh = ss.getSheetByName(SHEET_ORDERS);
  const rows = sh.getDataRange().getValues();
  const sold = { total: {}, byColor: {}, bySize: {} };
  for (let i = 1; i < rows.length; i++) {
    const no = String(rows[i][8] || '').trim();
    if (!no) continue;
    const color = String(rows[i][10] || '').trim();
    const size = String(rows[i][11] || '').trim();
    const qty = Number(rows[i][12]) || 0;
    sold.total[no] = (sold.total[no] || 0) + qty;
    sold.byColor[no + '|' + color] = (sold.byColor[no + '|' + color] || 0) + qty;
    sold.bySize[no + '|' + size] = (sold.bySize[no + '|' + size] || 0) + qty;
  }
  return sold;
}

// 내 주문 조회: 유튜브 닉네임 + 휴대전화 번호가 둘 다 일치하는 주문만 돌려줍니다.
// 개인정보 보호를 위해 주소/배송메모는 절대 내려주지 않습니다.
function lookupOrders(ytname, phone) {
  const digits = function (s) { return String(s || '').replace(/\D/g, ''); };
  const nick = String(ytname || '').trim();
  const ph = digits(phone);
  const out = { orders: [] };
  if (!nick || ph.length < 10) {
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  const rows = sh.getDataRange().getValues();
  // 열: 0주문시각 1유튜브닉네임 2수령인 3연락처 4주소 5배송메모 6배송지역 7입금자명
  //     8상품번호 9상품명 10칼라 11사이즈 12수량 13금액 14입금할총액 15입금확인
  let curStatus = '', curTotal = '';
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[14] !== '' && r[14] !== null) { // 주문의 첫 줄: 총액/상태가 적힌 줄
      curTotal = r[14];
      curStatus = String(r[15] || '대기');
    }
    if (String(r[1]).trim() === nick && digits(r[3]) === ph) {
      out.orders.push({
        date: String(r[0]),
        item: '[' + r[8] + '] ' + r[9],
        opt: r[10] + ' / ' + r[11] + ' / ' + r[12] + '개',
        amount: Number(r[13]) || 0,
        total: Number(curTotal) || 0,
        status: curStatus || '대기',
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// 고객이 주문서를 제출하면 재고를 확인한 뒤 [주문접수] 탭에 한 줄씩 기록합니다.
// 동시에 여러 명이 주문해도 재고가 초과되지 않도록 잠금(Lock)을 사용합니다.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 재고 확인 (상품 전체 재고 / 옵션별 재고 모두 검사)
    const prows = ss.getSheetByName(SHEET_PRODUCTS).getDataRange().getValues();
    const prodInfo = {};
    for (let i = 1; i < prows.length; i++) {
      const no = String(prows[i][0] || '').trim();
      if (!no) continue;
      const colors = String(prows[i][3] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      const sizes = String(prows[i][4] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      prodInfo[no] = { colors: colors, sizes: sizes, st: parseStock(prows[i][5], colors, sizes) };
    }
    const sold = soldByProduct(ss);
    const wantTotal = {}, wantOpt = {};
    data.items.forEach(function (it) {
      const no = String(it.no).trim();
      const qty = Number(it.qty) || 0;
      wantTotal[no] = (wantTotal[no] || 0) + qty;
      wantOpt[no + '|C|' + String(it.color).trim()] = (wantOpt[no + '|C|' + String(it.color).trim()] || 0) + qty;
      wantOpt[no + '|S|' + String(it.size).trim()] = (wantOpt[no + '|S|' + String(it.size).trim()] || 0) + qty;
    });
    for (const no in wantTotal) {
      const info = prodInfo[no];
      if (!info || !info.st) continue;
      const st = info.st;
      if (st.by === 'total') {
        const remain = Math.max(0, st.total - (sold.total[no] || 0));
        if (wantTotal[no] > remain) {
          return ContentService.createTextOutput(JSON.stringify({
            ok: false, error: 'stock', no: no, remaining: remain,
          })).setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        const opts = st.by === 'color' ? info.colors : info.sizes;
        const soldMap = st.by === 'color' ? sold.byColor : sold.bySize;
        const dim = st.by === 'color' ? 'C' : 'S';
        for (let j = 0; j < opts.length; j++) {
          const wanted = wantOpt[no + '|' + dim + '|' + opts[j]] || 0;
          if (!wanted) continue;
          const remain = Math.max(0, st.list[j] - (soldMap[no + '|' + opts[j]] || 0));
          if (wanted > remain) {
            return ContentService.createTextOutput(JSON.stringify({
              ok: false, error: 'stock', no: no, option: opts[j], remaining: remain,
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }

    const sh = ss.getSheetByName(SHEET_ORDERS);
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    data.items.forEach(function (it, idx) {
      sh.appendRow([
        now,
        data.ytname, data.name, data.phone, data.address,
        data.memo || '', data.region || '일반',
        data.payer || data.ytname,
        it.no, it.name, it.color, it.size, it.qty, it.price * it.qty,
        idx === 0 ? data.total : '',
        idx === 0 ? '대기' : '',
      ]);
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
