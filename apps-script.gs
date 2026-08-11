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
    p.appendRow(['상품번호', '상품명', '가격', '칼라(쉼표로 구분)', '사이즈(쉼표로 구분)']);
    p.appendRow(['101', '니트 가디건', 29000, '아이보리,블랙,핑크', 'Free']);
    p.appendRow(['102', '와이드 팬츠', 24000, '베이지,차콜', 'S,M,L']);
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

// 주문서 페이지가 열릴 때 상품 목록을 내려줍니다.
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  const products = {};
  for (let i = 1; i < rows.length; i++) {
    const [no, name, price, colors, sizes] = rows[i];
    if (!no || !name) continue;
    products[String(no).trim()] = {
      name: String(name).trim(),
      price: Number(price) || 0,
      colors: String(colors || '').split(',').map(s => s.trim()).filter(Boolean),
      sizes: String(sizes || '').split(',').map(s => s.trim()).filter(Boolean),
    };
  }
  return ContentService.createTextOutput(JSON.stringify(products))
    .setMimeType(ContentService.MimeType.JSON);
}

// 고객이 주문서를 제출하면 [주문접수] 탭에 한 줄씩 기록합니다.
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
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
}
