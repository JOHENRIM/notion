const pptxgen = require("pptxgenjs");

const RED = "DC2626";
const NAVY = "1E3A8A";
const INK = "1F2937";
const GRAY = "6B7280";
const LINE = "E5E7EB";
const F = "Malgun Gothic";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "Aurora World";
pres.title = "Amazon 봉제완구 랭킹 4대 지표 교차분석";

const W = 13.3, H = 7.5, M = 0.7;
const CW = W - M * 2; // 11.9

// ---------- helpers ----------
function pageTitle(slide, text, kicker) {
  if (kicker) {
    slide.addText(kicker, {
      x: M, y: 0.42, w: CW, h: 0.34, margin: 0,
      fontFace: F, fontSize: 16, bold: true, color: NAVY, charSpacing: 2,
    });
  }
  slide.addText(text, {
    x: M, y: kicker ? 0.76 : 0.5, w: CW, h: 0.8, margin: 0,
    fontFace: F, fontSize: 36, bold: true, color: RED, valign: "top",
  });
}

function pageNum(slide, n) {
  slide.addText(String(n), {
    x: W - M - 0.7, y: H - 0.62, w: 0.7, h: 0.32, margin: 0,
    fontFace: F, fontSize: 14, color: GRAY, align: "right",
  });
}

function caption(slide, text) {
  slide.addText(text, {
    x: M, y: H - 0.66, w: CW - 1.0, h: 0.4, margin: 0,
    fontFace: F, fontSize: 13, color: GRAY, valign: "top",
  });
}

// outlined card (the deck's line motif)
function card(slide, x, y, w, h, color) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: color || LINE, width: 1 },
  });
}

function badge(slide, x, y, d, label, color) {
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: color || RED }, line: { color: color || RED, width: 1 },
  });
  slide.addText(label, {
    x, y, w: d, h: d, margin: 0,
    fontFace: F, fontSize: 20, bold: true, color: "FFFFFF", align: "center", valign: "middle",
  });
}

// =====================================================================
// 1. TITLE
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };

  s.addText("AURORA  |  PRODUCT PLANNING", {
    x: M, y: 0.85, w: CW, h: 0.36, margin: 0,
    fontFace: F, fontSize: 16, bold: true, color: NAVY, charSpacing: 3,
  });

  s.addText("Amazon 봉제완구 랭킹\n4대 지표 교차분석", {
    x: M, y: 1.35, w: CW, h: 1.9, margin: 0,
    fontFace: F, fontSize: 46, bold: true, color: RED, lineSpacing: 56, valign: "top",
  });

  s.addText("지금 팔리는 상품이 아니라, 앞으로 팔릴 상품을 찾는 프레임", {
    x: M, y: 3.35, w: CW, h: 0.5, margin: 0,
    fontFace: F, fontSize: 26, color: INK,
  });

  const tags = ["BEST SELLERS", "MOVERS & SHAKERS", "NEW RELEASES", "MOST WISHED FOR"];
  const tw = 2.75, gap = (CW - tw * 4) / 3;
  tags.forEach((t, i) => {
    const x = M + i * (tw + gap);
    card(s, x, 4.35, tw, 0.95, i === 0 ? RED : NAVY);
    s.addText(t, {
      x: x + 0.1, y: 4.35, w: tw - 0.2, h: 0.95, margin: 0,
      fontFace: F, fontSize: 15, bold: true, color: i === 0 ? RED : NAVY,
      align: "center", valign: "middle",
    });
  });

  s.addText("미국 Amazon  >  Toys & Games  >  Stuffed Animals & Plush Toys", {
    x: M, y: 5.85, w: CW, h: 0.4, margin: 0,
    fontFace: F, fontSize: 20, bold: true, color: INK,
  });
  s.addText("작성 : 상품기획팀   |   기준일 : 2026. 08. 11.   |   랭킹은 시간·지역에 따라 변동", {
    x: M, y: 6.3, w: CW, h: 0.4, margin: 0,
    fontFace: F, fontSize: 15, color: GRAY,
  });

  s.addNotes("미국 아마존 봉제완구 카테고리의 4개 공식 랭킹(Best Sellers / Movers & Shakers / New Releases / Most Wished For)을 교차 분석하여, 오로라 상품기획의 우선 개발 방향을 도출한 보고서입니다.");
}

// =====================================================================
// 2. EXECUTIVE SUMMARY
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "핵심 요약", "EXECUTIVE SUMMARY");

  const rows = [
    ["4개 랭킹은 서로 다른 시점을 본다",
     "현재 매출 · 24시간 급등 · 신상품 반응 · 미래 구매의향"],
    ["교차 지점에서 성장 상품군이 드러난다",
     "미니 플러시·백참, 블라인드박스, 기능성 힐링 plush"],
    ["강점은 지키고, 키덜트·라이프스타일로 확장한다",
     "동물 리얼리즘·베이비 경쟁력 유지 + 수집·착용형 신규 라인"],
  ];

  let y = 1.85;
  rows.forEach((r, i) => {
    card(s, M, y, CW, 1.3, LINE);
    badge(s, M + 0.32, y + 0.35, 0.6, String(i + 1), RED);
    s.addText(r[0], {
      x: M + 1.1, y: y + 0.16, w: CW - 1.5, h: 0.52, margin: 0,
      fontFace: F, fontSize: 27, bold: true, color: NAVY, valign: "middle",
    });
    s.addText(r[1], {
      x: M + 1.1, y: y + 0.66, w: CW - 1.5, h: 0.52, margin: 0,
      fontFace: F, fontSize: 26, color: INK, valign: "middle",
    });
    y += 1.46;
  });

  s.addText([
    { text: "글로벌 봉제완구 시장 ", options: { fontSize: 22, color: INK } },
    { text: "$13.7B (2025) → $25.9B (2033)", options: { fontSize: 26, bold: true, color: RED } },
    { text: "  ·  연평균 약 8%대 성장", options: { fontSize: 22, color: INK } },
  ], { x: M, y: 6.4, w: CW, h: 0.5, margin: 0, fontFace: F });

  pageNum(s, 2);
  s.addNotes("보고의 결론을 한 장으로 압축. 4대 랭킹은 대체재가 아니라 시간축이 다른 지표이며, 겹치는 지점을 봐야 개발 우선순위가 나옵니다.");
}

// =====================================================================
// 3. 4대 지표 정의
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "4대 랭킹 지표, 무엇을 말하는가", "FRAMEWORK");

  const items = [
    ["BEST SELLERS", "현재 판매 검증", "지금 실제로 팔리는 상품\n시간 단위로 갱신", RED],
    ["MOVERS & SHAKERS", "단기 바이럴", "최근 24시간 순위 상승폭\nSNS·프로모션 반응", RED],
    ["NEW RELEASES", "신상품 초기 반응", "출시·출시예정 신상품 중\n판매 반응이 빠른 상품", NAVY],
    ["MOST WISHED FOR", "미래 구매 욕망", "위시리스트·레지스트리\n저장이 가장 많은 상품", NAVY],
  ];

  const cw = (CW - 0.4) / 2, ch = 2.15;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + 0.4);
    const y = 1.88 + Math.floor(i / 2) * (ch + 0.3);
    card(s, x, y, cw, ch, it[3]);
    s.addText(it[0], {
      x: x + 0.32, y: y + 0.18, w: cw - 0.64, h: 0.34, margin: 0,
      fontFace: F, fontSize: 17, bold: true, color: it[3], charSpacing: 1,
    });
    s.addText(it[1], {
      x: x + 0.32, y: y + 0.5, w: cw - 0.64, h: 0.58, margin: 0,
      fontFace: F, fontSize: 30, bold: true, color: INK, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.32, y: y + 1.1, w: cw - 0.64, h: 0.95, margin: 0,
      fontFace: F, fontSize: 26, color: INK, lineSpacing: 34, valign: "top",
    });
  });

  caption(s, "※ Best Sellers Rank는 수시 갱신, Movers & Shakers는 최근 24시간 기준 → 접속 시점·지역·재고에 따라 순위 상이");
  pageNum(s, 3);
  s.addNotes("Best Sellers = 현재, Movers = 단기, New Releases = 신상품, Most Wished For = 미래. 이 네 축을 분리해서 봐야 합니다.");
}

// =====================================================================
// 4. 교차분석 프레임
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "교차하면 기획 판단이 나온다", "CROSS ANALYSIS");

  const rows = [
    ["BEST + WISHED", "안정 SKU · 기본 라인업화, 재고 확대", RED],
    ["BEST + NEW", "차별화 소재·동물·사이즈로 후속 개발", RED],
    ["MOVERS + WISHED", "숏폼·인플루언서·팝업 테스트 상품", RED],
    ["NEW + WISHED", "6~12개월 선행 기획 후보", NAVY],
    ["MOVERS 단독", "일시적 프로모션 — 대량 개발 보류", GRAY],
    ["WISHED ↑ · BEST ↓", "가격·재고 장벽 → 프리미엄 선물형", NAVY],
  ];

  let y = 1.9;
  const rh = 0.68, gap = 0.11;
  rows.forEach((r) => {
    card(s, M, y, CW, rh, LINE);
    s.addShape(pres.ShapeType.ellipse, {
      x: M + 0.3, y: y + rh / 2 - 0.09, w: 0.18, h: 0.18,
      fill: { color: r[2] }, line: { color: r[2], width: 1 },
    });
    s.addText(r[0], {
      x: M + 0.68, y: y, w: 3.9, h: rh, margin: 0,
      fontFace: F, fontSize: 26, bold: true, color: r[2] === GRAY ? INK : r[2], valign: "middle",
    });
    s.addText("→", {
      x: M + 4.65, y: y, w: 0.5, h: rh, margin: 0,
      fontFace: F, fontSize: 22, color: GRAY, align: "center", valign: "middle",
    });
    s.addText(r[1], {
      x: M + 5.25, y: y, w: CW - 5.55, h: rh, margin: 0,
      fontFace: F, fontSize: 26, color: INK, valign: "middle",
    });
    y += rh + gap;
  });

  pageNum(s, 4);
  s.addNotes("한 지표만 보면 오판합니다. Movers 단독 급등은 할인 프로모션일 가능성이 높아 대량 개발 근거로 쓰면 안 됩니다.");
}

// =====================================================================
// 5. 성장 상품군 8
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "랭킹에서 읽히는 성장 상품군", "MARKET SIGNAL");

  const items = [
    "미니 플러시 · 클립온 · 백참",
    "블라인드박스형 봉제 굿즈",
    "힐링 · 수면 · 온열 기능성",
    "어글리 큐트 캐릭터",
    "리얼리즘 동물 plush",
    "베이비 센서리 plush",
    "푸드 · 디저트 plush",
    "외부 IP 콜라보 plush",
  ];

  const cw = (CW - 0.4) / 2, rh = 0.98, gap = 0.2;
  items.forEach((t, i) => {
    const x = M + (i % 2) * (cw + 0.4);
    const y = 1.95 + Math.floor(i / 2) * (rh + gap);
    const c = i < 4 ? RED : NAVY;
    card(s, x, y, cw, rh, c);
    badge(s, x + 0.26, y + 0.22, 0.54, String(i + 1), c);
    s.addText(t, {
      x: x + 0.98, y: y, w: cw - 1.25, h: rh, margin: 0,
      fontFace: F, fontSize: 26, bold: true, color: INK, valign: "middle",
    });
  });

  caption(s, "1~4 : 신규 확장 영역(키덜트·수집·감성 케어)   |   5~8 : 오로라 기존 강점 기반 심화 영역");
  pageNum(s, 5);
  s.addNotes("1 미니/클립온/백참 : Most Wished·New Releases 동시 진입 여부 확인. 2 블라인드박스 : Movers 급상승 여부. 3 기능성 : Best Seller 장기 잔존 여부. 4 어글리 큐트 : 위시리스트 저장률. 5 리얼리즘 : Best Seller 안정성. 6 베이비 : 리뷰 수. 7 푸드 : New Releases 반응. 8 IP : Best Seller와 Most Wished 중복 여부.");
}

// =====================================================================
// 6. 브랜드 분류 (정정)
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "브랜드 분류 정정", "BRAND MAPPING");

  const cw = (CW - 0.5) / 2;
  const own = [
    ["Palm Pals", "미니 · 백참"],
    ["Miyoni", "리얼리즘 동물"],
    ["Eco Nation", "친환경 · 재생소재"],
    ["ebba", "베이비 센서리"],
    ["YooHoo", "캐릭터 IP"],
  ];
  const comp = [
    ["Jellycat", "감성 · 프리미엄"],
    ["Squishmallows", "수집 · 대형 IP"],
    ["Douglas", "리얼리즘 동물"],
    ["GUND", "클래식 · 베이비"],
    ["Mary Meyer", "베이비 센서리"],
    ["POP MART", "블라인드박스"],
  ];

  [[own, "오로라월드 자사 브랜드", NAVY, 0], [comp, "벤치마크 · 경쟁 브랜드", RED, 1]].forEach(([list, head, color, col]) => {
    const x = M + col * (cw + 0.5);
    card(s, x, 1.88, cw, 4.4, color);
    s.addText(head, {
      x: x + 0.3, y: 2.03, w: cw - 0.6, h: 0.52, margin: 0,
      fontFace: F, fontSize: 27, bold: true, color: color, valign: "middle",
    });
    list.forEach((b, i) => {
      const y = 2.66 + i * 0.6;
      s.addText([
        { text: b[0], options: { bold: true, color: INK } },
        { text: "   " + b[1], options: { bold: false, color: GRAY } },
      ], {
        x: x + 0.3, y, w: cw - 0.6, h: 0.55, margin: 0,
        fontFace: F, fontSize: 26, valign: "middle",
      });
    });
  });

  caption(s, "※ 정정 : Mary Meyer · Douglas · GUND는 오로라 브랜드가 아닌 경쟁사입니다. 기존 자료의 혼용 분류를 위와 같이 분리했습니다.");
  pageNum(s, 6);
  s.addNotes("기존 초안에서 Mary Meyer가 ebba와 같은 계열처럼 표기되어 있었습니다. 자사 브랜드와 벤치마크 대상을 명확히 분리했습니다.");
}

// =====================================================================
// 7. 시장 전망
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "시장은 성인 컬렉터로 커진다", "MARKET SIZE");

  s.addChart(
    pres.ChartType.bar,
    [{ name: "시장 규모 (US$ B)", labels: ["2025", "2029(E)", "2033(E)"], values: [13.68, 19.0, 25.94] }],
    {
      x: M - 0.15, y: 1.9, w: 6.6, h: 4.4,
      barDir: "col", barGapWidthPct: 60,
      chartColors: [NAVY, NAVY, RED],
      varyColors: true,
      showTitle: false,
      showLegend: false,
      showValue: true, dataLabelPosition: "outEnd",
      dataLabelFontFace: F, dataLabelFontSize: 16, dataLabelColor: INK, dataLabelFormatCode: '$#,##0.0"B"',
      catAxisLabelFontFace: F, catAxisLabelFontSize: 16, catAxisLabelColor: INK,
      valAxisHidden: true,
      valGridLine: { style: "none" },
      catGridLine: { style: "none" },
      valAxisMaxVal: 30,
    }
  );

  const pts = [
    "성인 컬렉터가 성장 동력",
    "완구 → 자기표현 아이템",
    "수집 · 착용 · 공유형이 유리",
  ];
  let y = 1.95;
  pts.forEach((t) => {
    card(s, M + 6.85, y, CW - 6.85, 1.0, LINE);
    s.addText(t, {
      x: M + 7.15, y, w: CW - 7.5, h: 1.0, margin: 0,
      fontFace: F, fontSize: 26, color: INK, valign: "middle",
    });
    y += 1.15;
  });

  card(s, M + 6.85, 5.4, CW - 6.85, 0.9, RED);
  s.addText("CAGR 약 8.3%", {
    x: M + 6.85, y: 5.4, w: CW - 6.85, h: 0.9, margin: 0,
    fontFace: F, fontSize: 28, bold: true, color: RED, align: "center", valign: "middle",
  });

  caption(s, "출처 : Grand View Research, Stuffed Animals & Plush Toys Market (2025 → 2033 전망) · 2029년은 보간 추정치");
  pageNum(s, 7);
  s.addNotes("시장 자체가 아동 완구에서 성인 감성 소비로 확장되고 있어, 오로라의 신규 라인 확장 근거가 됩니다.");
}

// =====================================================================
// 8. 개발 우선순위
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "오로라 개발 우선순위", "ACTION PRIORITY");

  const rows = [
    ["4~6인치 미니 플러시 백참 · 클립온", "착용형 수집 트렌드"],
    ["블라인드박스형 캐릭터 봉제 시리즈", "언박싱 · 수집 문화"],
    ["온열 · 무게감 · 향기 기능성 plush", "힐링 · 수면 케어 확장"],
    ["어글리 큐트 신규 캐릭터", "키덜트 · SNS 바이럴"],
    ["Miyoni · Eco Nation 프리미엄 라인", "기존 강점 + 친환경"],
    ["ebba 베이비 센서리 강화", "출산 · 육아 선물 수요"],
    ["푸드 · 디저트 · 홈데코 plush", "라이프스타일 채널"],
    ["외부 IP 미니 plush 협업", "선물 · 위시리스트 수요"],
  ];

  let y = 1.85;
  const rh = 0.52, gap = 0.05;
  rows.forEach((r, i) => {
    const c = i < 3 ? RED : NAVY;
    s.addShape(pres.ShapeType.ellipse, {
      x: M, y: y + 0.06, w: 0.4, h: 0.4,
      fill: { color: i < 3 ? RED : "FFFFFF" }, line: { color: c, width: 1 },
    });
    s.addText(String(i + 1), {
      x: M, y: y + 0.06, w: 0.4, h: 0.4, margin: 0,
      fontFace: F, fontSize: 16, bold: true, color: i < 3 ? "FFFFFF" : c,
      align: "center", valign: "middle",
    });
    s.addText(r[0], {
      x: M + 0.68, y, w: 7.2, h: rh, margin: 0,
      fontFace: F, fontSize: 26, bold: true, color: INK, valign: "middle",
    });
    s.addText(r[1], {
      x: M + 7.95, y, w: CW - 7.95, h: rh, margin: 0,
      fontFace: F, fontSize: 26, color: c, valign: "middle",
    });
    if (i < rows.length - 1) {
      s.addShape(pres.ShapeType.line, {
        x: M, y: y + rh + gap / 2, w: CW, h: 0,
        line: { color: LINE, width: 1 },
      });
    }
    y += rh + gap;
  });

  caption(s, "1~3순위 : 2026 하반기 선행 개발 착수   |   4~8순위 : 2027 라인업 검토");
  pageNum(s, 8);
  s.addNotes("1~3순위를 우선 착수 대상으로 제안합니다. 나머지는 기존 라인 강화 및 후속 검토 대상입니다.");
}

// =====================================================================
// 9. 결론 & NEXT STEP
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  pageTitle(s, "결론 및 다음 단계", "CONCLUSION");

  card(s, M, 1.9, CW, 1.6, RED);
  s.addText("네 지표를 교차하면, 지금 잘 팔리는 봉제완구가 아니라\n앞으로 성장할 상품군이 보인다.", {
    x: M + 0.45, y: 1.9, w: CW - 0.9, h: 1.6, margin: 0,
    fontFace: F, fontSize: 28, bold: true, color: RED, lineSpacing: 40, valign: "middle",
  });

  const steps = [
    ["01", "랭킹 주 1회\n캡처 · 아카이브", "캡처 일시 병기 필수"],
    ["02", "교차 상위\nSKU 리스트업", "경쟁사 · 소재 · 가격대 분석"],
    ["03", "미니 · 백참 라인\n선행 기획", "2026 하반기 착수"],
  ];
  const cw = (CW - 0.6) / 3;
  steps.forEach((st, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 3.85, cw, 2.35, i === 0 ? RED : NAVY);
    s.addText(st[0], {
      x: x + 0.3, y: 4.02, w: cw - 0.6, h: 0.42, margin: 0,
      fontFace: F, fontSize: 20, bold: true, color: i === 0 ? RED : NAVY,
    });
    s.addText(st[1], {
      x: x + 0.3, y: 4.5, w: cw - 0.6, h: 1.1, margin: 0,
      fontFace: F, fontSize: 26, bold: true, color: INK, lineSpacing: 34, valign: "top",
    });
    s.addText(st[2], {
      x: x + 0.3, y: 5.65, w: cw - 0.6, h: 0.4, margin: 0,
      fontFace: F, fontSize: 16, color: GRAY,
    });
  });

  caption(s, "※ Amazon 랭킹은 접속 시점 · 지역 · 계정 · 재고에 따라 변동됩니다. 내부 보고 시 반드시 캡처 일시를 함께 기재하십시오.");
  pageNum(s, 9);
  s.addNotes("실행은 세 가지입니다. 랭킹 정기 캡처 체계화, 교차 상위 SKU 리스트업, 미니·백참 라인 선행 기획 착수.");
}

pres.writeFile({ fileName: "Aurora_Amazon_Plush_Ranking_Analysis.pptx" })
  .then((f) => console.log("saved:", f));
