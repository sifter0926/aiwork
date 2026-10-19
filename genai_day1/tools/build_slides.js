// 종합 예제(엑셀 상관·회귀·시각화 보고서) 슬라이드 생성 — 원본 덱 스타일(Arial/맑은 고딕 대체, 남색·주황) 재현
const pptxgen = require("pptxgenjs");
const fs = require("fs");

const J = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const data = J.rows; data.__corr = J.corr;
const OUT = process.argv[3];

const C = {
  navy: "12304D", navy2: "1E4468", big: "1C3F62", orange: "F2A65A", brown: "A64F16",
  teal: "1F6F78", cream: "F5F4EF", sky: "E8EEF2", card: "FCFCFA", border: "D9DDE1",
  body: "3D4A57", muted: "5E6977", pale: "BFD0DE", alt: "ECEEEA", grid: "C9C9C7",
};
const FONT = "Arial";
const FOOT = "생성형 AI 실무 활용 · 1일차";
const FOOT_DATA = FOOT + " · 강사 시연용 가상 데이터";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

const ls = (sz) => Math.round(sz * 1.4 * 10) / 10;
function T(slide, text, o) {
  const sz = o.fontSize || 13;
  slide.addText(text, Object.assign({
    fontFace: FONT, fontSize: sz, color: C.body, margin: 0, valign: "top",
    lineSpacing: ls(sz), isTextBox: true, fit: "none",
  }, o));
}
function label(s, text, y = 0.89) { T(s, text, { x: 0.89, y, w: 11.9, h: 0.3, fontSize: 14, bold: true, color: C.brown }); }
function title(s, text, { y = 1.24, size = 32, w = 11.9 } = {}) {
  T(s, text, { x: 0.89, y, w, h: size >= 32 ? 0.54 : 0.47, fontSize: size, bold: true, color: C.navy, objectName: "Title" });
}
function footer(s, text = FOOT) { T(s, text, { x: 0.89, y: 6.82, w: 11.9, h: 0.26, fontSize: 12, color: C.muted }); }
function badge(s, text, w) {
  s.addText(text, { x: 0.89, y: 0.89, w, h: 0.38, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.19,
    fill: { color: C.teal }, fontFace: FONT, fontSize: 14, bold: true, color: C.cream, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function card(s, x, y, w, h, { fill = C.card, line = C.border, r = 0.12 } = {}) {
  const o = { x, y, w, h, fill: { color: fill }, rectRadius: r };
  if (line) o.line = { color: line, width: 0.75 }; else o.line = { type: "none" };
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, o);
}
const run = (text, o = {}) => ({ text, options: Object.assign({ fontFace: FONT }, o) });

function table(s, rows, o) {
  // rows: array of arrays of strings or {text, options}; first row header
  const sz = o.fontSize || 12;
  const body = rows.map((r, i) => r.map((cell, j) => {
    const base = typeof cell === "string" ? { text: cell, options: {} } : cell;
    const head = i === 0 && !o.noHeader;
    const opt = Object.assign({
      fontFace: FONT, fontSize: sz, color: head ? C.cream : C.navy, bold: head,
      fill: { color: head ? C.navy : (i % 2 === 1 ? C.card : C.alt) },
      border: [{ type: "solid", pt: 1, color: C.grid }, { type: "solid", pt: 1, color: C.grid }, { type: "solid", pt: 1, color: C.grid }, { type: "solid", pt: 1, color: C.grid }],
      margin: [0.04, 0.07, 0.04, 0.07], valign: "middle",
      align: (o.right || []).includes(j) ? "right" : (o.center || []).includes(j) ? "center" : "left",
    }, base.options);
    return { text: base.text, options: opt };
  }));
  s.addTable(body, { x: o.x, y: o.y, w: o.w, colW: o.colW, rowH: o.rowH, autoPage: false });
}

// ---------- 공통 수치 ----------
const N = {
  r_ti: "0.88", r_rv: "−0.64", r_tv: "0.18", r_vs: "0.98", slope: "15.04", inter: "−209.6", r2: "0.78", pred: "211",
};
const temps = data.map((d) => d["최고기온"]);
const ice = data.map((d) => d["아이스음료"]);

// =====================================================================
// 1. 섹션 표지
let s = pres.addSlide(); s.background = { color: C.navy };
T(s, "09", { x: 7.44, y: 1.39, w: 5.01, h: 2.81, fontSize: 200, bold: true, color: C.big, align: "right", lineSpacing: 200 });
T(s, "심화 · 종합 예제 · 80분", { x: 0.89, y: 2.46, w: 11.9, h: 0.34, fontSize: 16, bold: true, color: C.orange });
T(s, "엑셀 데이터 분석 보고서", { x: 0.89, y: 3.02, w: 11.9, h: 0.76, fontSize: 48, bold: true, color: C.cream, objectName: "Title" });
T(s, "상관분석 · 회귀분석 · 시각화로 근거 있는 보고서를 만듭니다", { x: 0.89, y: 4.0, w: 11.9, h: 0.38, fontSize: 18, color: C.pale });
let cx = 0.89;
[["상황과 데이터", 1.72], ["상관분석", 1.17], ["회귀분석", 1.17], ["시각화", 0.97], ["보고서 작성", 1.53]].forEach(([t, w], i, a) => {
  s.addText(t, { x: cx, y: 4.6, w, h: 0.44, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.22,
    fill: { color: i === a.length - 1 ? C.teal : C.navy2 }, fontFace: FONT, fontSize: 14, color: C.cream, align: "center", valign: "middle", margin: 0, isTextBox: true });
  cx += w + 0.11;
});
s.addNotes("8교시 종합 실습을 마친 뒤 이어서 하는 심화 예제입니다. 6교시(엑셀 수식)와 7교시(검증), 8교시(보고서)를 데이터 분석으로 확장합니다. 계산은 엑셀이, 문장 초안은 AI가, 해석과 검증은 사람이 한다는 역할 분담을 다시 강조해 주세요. 실습 파일: cafe_daily_202609.xlsx(데이터 시트만 있는 실습용 / 정답 시트 포함 강사용).");

// =====================================================================
// 2. 미션
s = pres.addSlide(); s.background = { color: C.sky };
badge(s, "종합 예제 · 미션", 1.73);
title(s, "9월 판매 데이터로 분석 보고서 만들기", { y: 1.38 });
card(s, 0.89, 2.2, 11.56, 1.17, { fill: C.navy, line: null });
T(s, [run("상황", { bold: true, color: C.orange }), run("　나는 베이커리 카페 본사 운영지원팀 직원입니다. 해운대점 9월 일별 데이터(30일)를 엑셀로 분석해 ‘무엇이 판매를 움직이는지’ 근거를 담은 1쪽 보고서를 팀장님께 드려야 합니다.", { color: C.cream })],
  { x: 1.16, y: 2.42, w: 11.05, h: 0.8, fontSize: 15 });
const steps = [
  ["STEP 1", "데이터 확인", "열 이름·단위 확인, 빈칸과 이상한 값 찾기"],
  ["STEP 2", "상관분석", "CORREL, 상관 분석 도구로 함께 움직이는 변수 찾기"],
  ["STEP 3", "회귀분석", "SLOPE, 회귀 분석 도구로 영향의 크기와 예측값 구하기"],
  ["STEP 4", "시각화", "산점도+추세선, 콤보 차트로 한눈에 보이기"],
  ["STEP 5", "보고서", "AI로 초안 쓰고 수치·해석 검증해 완성"],
];
const sw = (11.56 - 4 * 0.17) / 5;
steps.forEach(([k, h, d], i) => {
  const x = 0.89 + i * (sw + 0.17), last = i === 4;
  card(s, x, 3.67, sw, 1.95, { fill: last ? C.navy : C.card, line: null, r: 0.1 });
  T(s, k, { x: x + 0.22, y: 3.9, w: sw - 0.4, h: 0.3, fontSize: 14, bold: true, color: last ? C.orange : C.brown });
  T(s, h, { x: x + 0.22, y: 4.25, w: sw - 0.4, h: 0.33, fontSize: 18, bold: true, color: last ? C.cream : C.navy });
  T(s, d, { x: x + 0.22, y: 4.66, w: sw - 0.4, h: 0.85, fontSize: 13, color: last ? C.pale : C.body });
});
T(s, [run("산출물", { bold: true, color: C.navy }), run("　상관계수 표 · 회귀식과 예측값 · 차트 2개 · 1쪽 분석 보고서", { color: C.body })],
  { x: 0.89, y: 5.88, w: 11.56, h: 0.3, fontSize: 14 });
footer(s);
s.addNotes("진행 순서를 먼저 보여 주고 시작합니다. STEP 1~4는 엑셀로 직접 하고, STEP 5에서만 생성형 AI를 씁니다. AI에게 원본 데이터를 통째로 주지 않고 엑셀로 구한 결과 수치만 준다는 점(7교시 보안)을 미리 예고해 주세요.");

// =====================================================================
// 3. 데이터
s = pres.addSlide(); s.background = { color: C.sky };
label(s, "종합 예제 ① 상황과 데이터", 0.78);
title(s, "해운대점 9월 일별 판매 데이터 (30일)", { y: 1.12, size: 26 });
const head = ["날짜", "요일", "주말", "최고기온\n(℃)", "강수량\n(mm)", "방문객수\n(명)", "아이스음료\n(잔)", "따뜻한음료\n(잔)", "매출액\n(천원)"];
const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I"].map((t) => ({ text: t, options: { fill: { color: "DDE3E8" }, color: C.muted, fontSize: 10, bold: true, align: "center" } }));
const rowsD = data.slice(0, 7).map((d) => [d["날짜"].slice(5), d["요일"], String(d["주말"]), d["최고기온"].toFixed(1), String(d["강수량"]), String(d["방문객수"]), String(d["아이스음료"]), String(d["따뜻한음료"]), d["매출액"].toLocaleString("en-US")]);
const tbl = [letters, head.map((h) => ({ text: h, options: { fill: { color: C.navy }, color: C.cream, bold: true, align: "center" } }))]
  .concat(rowsD.map((r, i) => r.map((c, j) => ({ text: c, options: { fill: { color: i % 2 === 0 ? C.card : C.alt }, align: j <= 2 ? "center" : "right" } }))))
  .concat([[{ text: "⋮  9월 30일까지 같은 형식으로 30행 (2~31행)", options: { colspan: 9, align: "center", color: C.muted, fill: { color: C.card } } }]]);
table(s, tbl, { x: 0.89, y: 1.78, w: 8.25, colW: [0.78, 0.62, 0.62, 0.95, 0.9, 0.95, 1.0, 1.0, 0.93], rowH: [0.26, 0.52, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33], fontSize: 12, noHeader: true });
card(s, 9.4, 1.78, 3.05, 4.55, { fill: C.navy, line: null, r: 0.12 });
T(s, "분석 질문", { x: 9.65, y: 1.98, w: 2.6, h: 0.3, fontSize: 15, bold: true, color: C.orange });
const qs = [
  ["Q1", "아이스음료 판매와 함께 움직이는 것은?", "상관분석"],
  ["Q2", "기온이 1℃ 오르면 몇 잔 더 팔릴까?", "회귀분석"],
  ["Q3", "방문객 수는 비·주말·기온 중 무엇의 영향이 클까?", "다중회귀"],
  ["Q4", "팀장님이 한눈에 볼 차트는?", "시각화"],
];
qs.forEach(([q, t, m], i) => {
  const y = 2.42 + i * 0.95;
  T(s, [run(q + "  ", { bold: true, color: C.orange }), run(t, { color: C.cream })], { x: 9.65, y, w: 2.62, h: 0.6, fontSize: 13 });
  T(s, "→ " + m, { x: 9.65, y: y + 0.6, w: 2.62, h: 0.26, fontSize: 12, bold: true, color: C.pale });
});
T(s, [run("열 설명", { bold: true, color: C.navy }), run("　주말: 토·일 = 1, 평일 = 0 · 매출액 단위는 천 원 · 실습 파일 ", { color: C.body }), run("cafe_daily_202609.xlsx", { bold: true, color: C.teal }), run(" 의 ‘데이터’ 시트", { color: C.body })],
  { x: 0.89, y: 5.6, w: 8.3, h: 0.6, fontSize: 13 });
footer(s, FOOT_DATA);
s.addNotes("엑셀의 열 문자(A~I)를 함께 보여 주는 이유는 이후 수식에서 D열(최고기온), G열(아이스음료)처럼 열 위치로 말하기 때문입니다. 주말은 1/0 숫자로 바꿔 두어야 회귀분석에 넣을 수 있다는 점(더미 변수)을 짧게 언급해 주세요. 30행 전체는 실습 파일에 있습니다.");

// =====================================================================
// 4. 개념
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "상관분석 · 회귀분석");
title(s, "‘함께 움직이나’와 ‘얼마나 움직이나’");
[[0.89, "상관분석", "두 변수가 함께 움직이는 방향과 강도", [["숫자", "상관계수 r · −1 ~ +1"], ["엑셀", "CORREL · 데이터 분석 > 상관 분석"], ["질문", "더운 날 아이스음료도 많이 팔리나?"]]],
 [6.75, "회귀분석", "한 변수가 변할 때 다른 변수가 얼마나 변하는지 식으로 표현", [["숫자", "y = a + b·x · 결정계수 R² 0 ~ 1"], ["엑셀", "SLOPE·INTERCEPT·RSQ · 회귀 분석"], ["질문", "1℃ 오르면 몇 잔? 28℃면 몇 잔?"]]]].forEach(([x, h, d, items]) => {
  card(s, x, 2.06, 5.69, 2.3);
  T(s, h, { x: x + 0.25, y: 2.28, w: 5.2, h: 0.33, fontSize: 18, bold: true, color: C.navy });
  T(s, d, { x: x + 0.25, y: 2.66, w: 5.2, h: 0.3, fontSize: 13, color: C.body });
  items.forEach(([k, v], i) => T(s, [run(k + "  ", { bold: true, color: C.teal }), run(v, { color: C.navy })], { x: x + 0.25, y: 3.12 + i * 0.36, w: 5.2, h: 0.3, fontSize: 14 }));
});
const scale = [["−1 ~ −0.7", "강한 음(−)", "1E4468", C.cream], ["−0.7 ~ −0.3", "보통 음(−)", "7FA7BF", C.navy], ["−0.3 ~ 0.3", "거의 없음", "ECEEEA", C.navy], ["0.3 ~ 0.7", "보통 양(+)", "F2C08F", C.navy], ["0.7 ~ 1", "강한 양(+)", "C8692A", C.cream]];
T(s, "상관계수 r 읽는 법 (실무용 기준)", { x: 0.89, y: 4.6, w: 7, h: 0.3, fontSize: 13, bold: true, color: C.navy });
scale.forEach(([r, t, f, tc], i) => {
  const x = 0.89 + i * 1.47;
  s.addShape(pres.shapes.RECTANGLE, { x, y: 4.98, w: 1.43, h: 0.92, fill: { color: f }, line: { type: "none" } });
  T(s, r, { x: x + 0.1, y: 5.1, w: 1.25, h: 0.28, fontSize: 13, bold: true, color: tc, align: "center" });
  T(s, t, { x: x + 0.1, y: 5.45, w: 1.25, h: 0.28, fontSize: 13, color: tc, align: "center" });
});
card(s, 8.4, 4.6, 4.05, 1.3, { fill: C.navy, line: null });
T(s, "주의 · 상관은 원인이 아닙니다", { x: 8.62, y: 4.78, w: 3.65, h: 0.3, fontSize: 14, bold: true, color: C.orange });
T(s, "아이스크림 판매와 물놀이 사고는 함께 늘지만, 원인은 ‘더운 날씨’입니다.", { x: 8.62, y: 5.14, w: 3.65, h: 0.62, fontSize: 13, color: C.cream });
footer(s);
s.addNotes("상관분석은 방향과 강도만, 회귀분석은 ‘얼마나’와 ‘예측’까지 알려 준다고 구분합니다. r 기준(0.3/0.7)은 분야마다 다른 실무용 어림값이라고 덧붙여 주세요. 오른쪽 예시로 상관과 인과를 구분하는 습관을 먼저 잡아 두면 마지막 검증 단계에서 다시 활용할 수 있습니다.");

// =====================================================================
// 5. 엑셀 준비
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "엑셀 준비");
title(s, "‘데이터 분석’ 도구를 먼저 켜 둡니다");
const prep = [["①", "파일 > 옵션 > 추가 기능"], ["②", "아래 ‘관리: Excel 추가 기능’ → [이동]"], ["③", "‘분석 도구’ 체크 → [확인]"], ["④", "데이터 탭 오른쪽 끝 [데이터 분석] 확인"]];
prep.forEach(([n, t], i) => {
  const last = i === 3;
  card(s, 0.89, 2.06 + i * 0.68, 5.3, 0.56, { fill: last ? C.navy : C.card, line: last ? null : C.border, r: 0.1 });
  T(s, [run(n + "  ", { bold: true, color: last ? C.orange : C.teal }), run(t, { color: last ? C.cream : C.navy })], { x: 1.1, y: 2.06 + i * 0.68 + 0.14, w: 5.0, h: 0.3, fontSize: 14 });
});
T(s, "Mac은 도구 > Excel 추가 기능 > ‘분석 도구’ 체크", { x: 0.89, y: 4.85, w: 5.3, h: 0.28, fontSize: 12, color: C.muted });
table(s, [
  ["", "함수", "데이터 분석 도구"],
  [{ text: "예", options: { bold: true } }, "CORREL, SLOPE, RSQ", "상관 분석, 회귀 분석"],
  [{ text: "결과", options: { bold: true } }, "셀 하나에 값 하나", "표 전체를 한 번에"],
  [{ text: "장점", options: { bold: true } }, "자동 재계산", "P-값까지 자세한 통계"],
  [{ text: "언제", options: { bold: true } }, "핵심 수치 몇 개", "여러 변수 · 보고서 근거표"],
], { x: 6.55, y: 2.06, w: 5.9, colW: [0.8, 2.3, 2.8], rowH: 0.5, fontSize: 13 });
card(s, 0.89, 5.35, 11.56, 0.68, { fill: C.navy, line: null, r: 0.1 });
T(s, [run("오늘은 두 가지를 모두 씁니다 · ", { color: C.cream }), run("함수로 핵심 수치를 확인", { bold: true, color: C.orange }), run("하고, ", { color: C.cream }), run("분석 도구로 보고서용 근거표", { bold: true, color: C.orange }), run("를 만듭니다.", { color: C.cream })],
  { x: 1.12, y: 5.53, w: 11.1, h: 0.32, fontSize: 15 });
footer(s);
s.addNotes("분석 도구는 기본으로 꺼져 있어 첫 시간에 대부분 막힙니다. 화면을 공유하며 함께 켜 주세요. 분석 도구 결과는 값으로 붙여 넣어지므로 데이터를 고치면 다시 실행해야 하고, 함수는 자동으로 다시 계산된다는 차이를 짚어 주세요. 웹용 엑셀에는 분석 도구가 없으므로 함수로 진행합니다.");

// =====================================================================
// 6. 상관분석 방법
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "상관분석 · 엑셀로 하기");
title(s, "함수 한 줄 또는 도구 한 번이면 됩니다");
const cw = (11.56 - 2 * 0.17) / 3;
[["방법 1 · 함수", [run("=CORREL(D2:D31,G2:G31)", { bold: true, color: C.brown, fontSize: 15 })], ["최고기온(D)과 아이스음료(G)의", "상관계수 하나 → 0.88"]],
 ["방법 2 · 분석 도구", [run("데이터 > 데이터 분석 > 상관 분석", { bold: true, color: C.brown, fontSize: 14 })], ["입력 범위 D1:I31 · 첫째 행 이름표 ✓", "출력 범위 K1 → 6×6 상관표"]],
 ["방법 3 · 색 입히기", [run("홈 > 조건부 서식 > 색조", { bold: true, color: C.brown, fontSize: 14 })], ["상관표 숫자 칸을 선택하고 적용", "강한 관계가 진한 색으로 보임"]]].forEach(([h, code, d], i) => {
  const x = 0.89 + i * (cw + 0.17);
  card(s, x, 2.06, cw, 2.2);
  T(s, h, { x: x + 0.25, y: 2.28, w: cw - 0.45, h: 0.33, fontSize: 18, bold: true, color: C.navy });
  T(s, code, { x: x + 0.25, y: 2.78, w: cw - 0.45, h: 0.35, fontSize: 15 });
  T(s, d.join("\n"), { x: x + 0.25, y: 3.3, w: cw - 0.45, h: 0.7, fontSize: 13, color: C.body });
});
card(s, 0.89, 4.5, 11.56, 1.5, { fill: C.navy, line: null });
T(s, "AI에게 묻기 · 데이터 대신 열 이름과 위치만", { x: 1.15, y: 4.68, w: 11, h: 0.3, fontSize: 13, bold: true, color: C.orange });
T(s, "“엑셀에서 D열은 최고기온, G열은 아이스음료이고 데이터는 2~31행에 있어. 두 열의 상관계수를 구하는 수식과, 결과를 해석하는 기준을 알려 줘.”", { x: 1.15, y: 5.04, w: 11.05, h: 0.75, fontSize: 15, color: C.cream });
footer(s);
s.addNotes("CORREL의 두 범위는 순서를 바꿔도 결과가 같습니다(회귀의 SLOPE와 다른 점). 분석 도구의 입력 범위에는 숫자 열만 넣어야 하므로 날짜·요일(A~B)은 빼고, 주말(C)도 0/1이라 상관표에서는 제외했습니다. 출력된 상관표는 아래쪽 삼각형만 채워지는 것이 정상입니다. 6교시처럼 AI에게는 열 위치만 알려 줍니다.");

// =====================================================================
// 7. 상관분석 결과 (히트맵 표)
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "상관분석 결과");
title(s, "아이스음료는 기온, 방문객은 비와 함께 움직입니다", { size: 28 });
const vars = ["최고기온", "강수량", "방문객수", "아이스음료", "따뜻한음료", "매출액"];
const corr = data.__corr;
function heat(v) {
  if (v >= 0.999) return ["E3E6E9", C.muted];
  if (v >= 0.7) return ["C8692A", C.cream];
  if (v >= 0.3) return ["F2C08F", C.navy];
  if (v > -0.3) return ["FCFCFA", C.navy];
  if (v > -0.7) return ["7FA7BF", C.navy];
  return ["1E4468", C.cream];
}
const fmt = (v) => (v < 0 ? "−" : "") + Math.abs(v).toFixed(2);
const hrows = [[""].concat(vars).map((t) => ({ text: t, options: { fill: { color: C.navy }, color: C.cream, bold: true, align: "center", fontSize: 12 } }))];
vars.forEach((a, i) => {
  const r = [{ text: a, options: { fill: { color: C.navy }, color: C.cream, bold: true, fontSize: 12 } }];
  vars.forEach((b, j) => {
    if (j > i) r.push({ text: "", options: { fill: { color: C.card } } });
    else { const v = corr[a][b]; const [f, tc] = heat(v); r.push({ text: fmt(v), options: { fill: { color: f }, color: tc, align: "center", bold: (a === "아이스음료" && b === "최고기온") || (a === "방문객수" && b === "강수량") } }); }
  });
  hrows.push(r);
});
table(s, hrows, { x: 0.89, y: 2.0, w: 7.3, colW: [1.24, 1.01, 1.01, 1.01, 1.01, 1.01, 1.01], rowH: 0.5, fontSize: 13, noHeader: true });
T(s, "엑셀 ‘상관 분석’ 출력 + 조건부 서식(색조) · 아래쪽 삼각형만 채워지는 것이 정상", { x: 0.89, y: 5.62, w: 7.3, h: 0.28, fontSize: 12, color: C.muted });
const finds = [["0.88", "최고기온 ↔ 아이스음료", "강한 양(+) · 더운 날 아이스음료가 많이 팔림", C.brown],
  ["−0.64", "강수량 ↔ 방문객수", "보통 음(−) · 비 오는 날 손님이 줄어듦", C.teal],
  ["0.18", "최고기온 ↔ 방문객수", "거의 없음 · 기온과 손님 수는 무관", C.muted]];
finds.forEach(([v, h, d, col], i) => {
  const y = 2.0 + i * 1.22;
  card(s, 8.45, y, 4.0, 1.08);
  T(s, v, { x: 8.65, y: y + 0.16, w: 1.15, h: 0.45, fontSize: 24, bold: true, color: col });
  T(s, h, { x: 9.85, y: y + 0.18, w: 2.5, h: 0.3, fontSize: 14, bold: true, color: C.navy });
  T(s, d, { x: 8.65, y: y + 0.6, w: 3.7, h: 0.4, fontSize: 12, color: C.body });
});
T(s, [run("덤", { bold: true, color: C.navy }), run(" · 방문객수 ↔ 매출액 0.98 → 손님 수가 곧 매출", { color: C.body })], { x: 8.45, y: 5.62, w: 4.0, h: 0.28, fontSize: 12 });
footer(s, FOOT_DATA);
s.addNotes("진한 주황(0.7 이상)과 진한 파랑(−0.7 이하)부터 읽게 합니다. 핵심 발견: 기온–아이스음료 0.88, 강수량–방문객수 −0.64, 그리고 기온–방문객수 0.18처럼 ‘관계가 거의 없다’는 것도 중요한 결과입니다. 따뜻한음료와 기온은 −0.53으로 반대 방향이라는 점도 함께 찾아보게 하세요.");

// =====================================================================
// 8. 회귀분석 방법
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "회귀분석 · 엑셀로 하기");
title(s, "기온으로 아이스음료 판매량을 설명해 봅니다");
T(s, "함수로 핵심 수치 구하기", { x: 0.89, y: 2.0, w: 6.5, h: 0.3, fontSize: 14, bold: true, color: C.navy });
table(s, [
  ["구하는 것", "수식", "결과"],
  ["기울기 b", { text: "=SLOPE(G2:G31,D2:D31)", options: { color: C.brown, bold: true } }, "15.04"],
  ["절편 a", { text: "=INTERCEPT(G2:G31,D2:D31)", options: { color: C.brown, bold: true } }, "−209.6"],
  ["결정계수 R²", { text: "=RSQ(G2:G31,D2:D31)", options: { color: C.brown, bold: true } }, "0.78"],
  ["28℃일 때 예측", { text: "=FORECAST.LINEAR(28,G2:G31,D2:D31)", options: { color: C.brown, bold: true } }, "211"],
], { x: 0.89, y: 2.38, w: 6.6, colW: [1.45, 4.1, 1.05], rowH: 0.46, fontSize: 13, right: [2] });
T(s, [run("순서 주의", { bold: true, color: C.brown }), run(" · (y 범위, x 범위) — 결과(아이스음료 G)를 먼저 씁니다", { color: C.body })], { x: 0.89, y: 4.82, w: 6.6, h: 0.3, fontSize: 13 });
card(s, 7.75, 2.0, 4.7, 3.12);
T(s, "데이터 분석 > 회귀 분석", { x: 8.0, y: 2.2, w: 4.3, h: 0.33, fontSize: 17, bold: true, color: C.navy });
[["Y축 입력 범위", "G1:G31 (아이스음료)"], ["X축 입력 범위", "D1:D31 (최고기온)"], ["이름표", "✓ 첫 행이 열 이름"], ["출력 옵션", "새로운 워크시트"], ["잔차", "✓ 잔차 · ✓ 선적합도"]].forEach(([k, v], i) =>
  T(s, [run(k + "  ", { bold: true, color: C.teal }), run(v, { color: C.navy })], { x: 8.0, y: 2.68 + i * 0.44, w: 4.3, h: 0.3, fontSize: 13 }));
card(s, 0.89, 5.35, 11.56, 0.68, { fill: C.navy, line: null, r: 0.1 });
T(s, [run("회귀식 · ", { color: C.cream }), run("아이스음료(잔) = 15.04 × 최고기온(℃) − 209.6", { bold: true, color: C.orange }), run("　결정계수 R² = 0.78", { color: C.cream })],
  { x: 1.12, y: 5.53, w: 11.1, h: 0.32, fontSize: 15 });
footer(s, FOOT_DATA);
s.addNotes("SLOPE·INTERCEPT·RSQ·FORECAST.LINEAR는 모두 (y, x) 순서입니다. 순서를 바꾸면 기온을 아이스음료로 설명하는 엉뚱한 식이 나온다는 점을 실제로 바꿔 보여 주면 효과적입니다. FORECAST.LINEAR가 없는 구버전은 FORECAST를 씁니다. 회귀 분석 도구 결과는 다음 장에서 읽습니다.");

// =====================================================================
// 9. 회귀 결과 읽기
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "회귀분석 결과 읽기");
title(s, "출력표에서 세 숫자만 읽으면 됩니다");
const hi = { fill: { color: "F7D9B9" }, bold: true };
T(s, "회귀분석 통계량", { x: 0.89, y: 2.0, w: 3, h: 0.28, fontSize: 13, bold: true, color: C.navy });
table(s, [
  ["다중 상관계수", "0.88"], [{ text: "결정계수", options: hi }, { text: "0.78", options: hi }], ["조정된 결정계수", "0.77"], ["표준 오차", "17.5"], ["관측수", "30"],
], { x: 0.89, y: 2.32, w: 3.0, colW: [1.9, 1.1], rowH: 0.36, fontSize: 12, noHeader: true, right: [1] });
T(s, "계수표", { x: 4.1, y: 2.0, w: 3, h: 0.28, fontSize: 13, bold: true, color: C.navy });
table(s, [
  ["", "계수", "표준 오차", "t 통계량", "P-값"],
  ["Y 절편", "−209.59", "42.81", "−4.90", "3.69E-05"],
  [{ text: "최고기온(℃)", options: { bold: true } }, { text: "15.04", options: hi }, "1.52", "9.91", { text: "1.17E-10", options: hi }],
], { x: 4.1, y: 2.32, w: 4.35, colW: [1.2, 0.75, 0.8, 0.75, 0.85], rowH: 0.4, fontSize: 11, right: [1, 2, 3, 4] });
T(s, "엑셀 ‘회귀 분석’ 출력 중 일부 · 1.17E-10 = 0.000000000117", { x: 4.1, y: 3.62, w: 4.35, h: 0.28, fontSize: 11, color: C.muted });
const reads = [["①", "결정계수 0.78", "판매량 변동의 78%를 기온으로 설명"], ["②", "계수 15.04", "최고기온 1℃ ↑ → 아이스음료 약 15잔 ↑"], ["③", "P-값 < 0.05", "우연이라 보기 어려운 관계 (유의함)"]];
reads.forEach(([n, h, d], i) => {
  const y = 2.0 + i * 1.08;
  card(s, 8.7, y, 3.75, 0.96);
  T(s, [run(n + " ", { bold: true, color: C.teal }), run(h, { bold: true, color: C.navy })], { x: 8.9, y: y + 0.13, w: 3.4, h: 0.3, fontSize: 15 });
  T(s, d, { x: 8.9, y: y + 0.46, w: 3.45, h: 0.4, fontSize: 12, color: C.body });
});
card(s, 0.89, 4.2, 7.56, 0.9, { fill: C.card, line: C.border });
T(s, [run("예측하기", { bold: true, color: C.navy }), run("　내일 최고기온 28℃ 예보 → 15.04 × 28 − 209.6 ≈ ", { color: C.body }), run("211잔", { bold: true, color: C.brown }), run(" 준비\n표준 오차 17.5 → 실제 판매는 예측과 보통 ±18잔 정도 차이", { color: C.body })],
  { x: 1.12, y: 4.35, w: 7.2, h: 0.65, fontSize: 13 });
card(s, 0.89, 5.35, 11.56, 0.68, { fill: C.navy, line: null, r: 0.1 });
T(s, [run("믿을 수 있는 범위 · ", { color: C.cream }), run("관측한 기온 23.5~31.5℃ 안에서만", { bold: true, color: C.orange }), run(" 예측합니다. 40℃나 0℃처럼 범위 밖은 식이 맞지 않습니다.", { color: C.cream })],
  { x: 1.12, y: 5.53, w: 11.1, h: 0.32, fontSize: 15 });
footer(s, FOOT_DATA);
s.addNotes("엑셀 출력표에는 숫자가 많지만 보고서에 필요한 것은 결정계수, 계수, P-값 세 가지입니다. P-값은 엑셀에서 1.17E-10처럼 지수로 표시되므로 ‘0.05보다 아주 작다’고 읽으면 됩니다. 절편 −209.6은 ‘0℃면 −210잔’이라는 뜻이 되어 말이 안 되는데, 이것이 범위 밖 예측이 위험한 이유입니다.");

// =====================================================================
// 10. 다중회귀
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "한 걸음 더 · 다중회귀");
title(s, "방문객 수에는 무엇이 가장 크게 영향을 줄까");
T(s, [run("회귀 분석 설정", { bold: true, color: C.navy }), run("　Y축 F1:F31(방문객수) · X축 ", { color: C.body }), run("C1:E31", { bold: true, color: C.brown }), run(" (주말·최고기온·강수량 — 붙어 있는 세 열)", { color: C.body })],
  { x: 0.89, y: 2.0, w: 11.5, h: 0.3, fontSize: 13 });
table(s, [
  ["변수", "계수", "P-값", "해석"],
  ["주말 (1/0)", { text: "87.1", options: { bold: true } }, "< 0.001", "주말이면 약 87명 더 방문"],
  ["강수량 (mm)", { text: "−2.84", options: { bold: true } }, "< 0.001", "비 10mm당 약 28명 감소"],
  ["최고기온 (℃)", "2.65", "0.047", "1℃당 약 3명 · 영향은 작음"],
  ["Y 절편", "178.1", "", ""],
], { x: 0.89, y: 2.48, w: 7.3, colW: [1.6, 1.0, 1.1, 3.6], rowH: 0.46, fontSize: 13, right: [1, 2] });
T(s, "결정계수 R² 0.93 · 관측수 30 · 소수 셋째 자리 반올림", { x: 0.89, y: 4.88, w: 7.3, h: 0.28, fontSize: 12, color: C.muted });
card(s, 8.45, 2.48, 4.0, 2.65);
T(s, "방문객 수 설명력 (R²)", { x: 8.7, y: 2.68, w: 3.6, h: 0.3, fontSize: 14, bold: true, color: C.navy });
T(s, [run("0.03", { bold: true, color: C.muted, fontSize: 36 }), run("  →  ", { color: C.muted, fontSize: 20 }), run("0.93", { bold: true, color: C.brown, fontSize: 36 })], { x: 8.7, y: 3.1, w: 3.6, h: 0.75, fontSize: 36, lineSpacing: 44 });
T(s, "기온 하나만 넣었을 때 → 주말·비·기온을 함께 넣었을 때", { x: 8.7, y: 3.95, w: 3.6, h: 0.6, fontSize: 12, color: C.body });
card(s, 0.89, 5.35, 11.56, 0.68, { fill: C.navy, line: null, r: 0.1 });
T(s, [run("결론 · ", { color: C.cream }), run("손님 수는 주말과 비가 좌우하고, 기온은 손님이 ‘무엇을 사는지’를 바꿉니다.", { bold: true, color: C.orange })],
  { x: 1.12, y: 5.53, w: 11.1, h: 0.32, fontSize: 15 });
footer(s, FOOT_DATA);
s.addNotes("다중회귀는 X축 범위에 여러 열을 한 번에 넣으면 되는데, 열이 서로 붙어 있어야 합니다(그래서 주말·최고기온·강수량을 C~E에 배치). 함수로는 =LINEST(F2:F31,C2:E31,TRUE,TRUE)로 같은 계수를 구할 수 있지만 결과가 역순으로 나와 헷갈리므로 분석 도구를 권합니다. 상관분석의 기온–방문객 0.18과 이 결과가 서로 맞는다는 점을 연결해 주세요.");

// =====================================================================
// 11. 시각화 ① 산점도
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "시각화 ① 산점도와 추세선");
title(s, "관계는 산점도로, 추세선에 식과 R²를 표시합니다", { size: 28 });
function scatter(sl, x, y, w, h, { titleText = "최고기온과 아이스음료 판매량 (9월, 30일)", small = false } = {}) {
  sl.addChart(pres.charts.SCATTER, [{ name: "최고기온", values: temps }, { name: "아이스음료(잔)", values: ice }], {
    x, y, w, h, objectName: "IceScatter",
    lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: small ? 6 : 8, chartColors: [C.teal],
    showTitle: true, title: titleText, titleFontSize: small ? 11 : 13, titleColor: C.navy, titleFontFace: FONT,
    showLegend: false,
    valAxisMinVal: 100, valAxisMaxVal: 300, catAxisMinVal: 22, catAxisMaxVal: 33,
    showValAxisTitle: true, valAxisTitle: "아이스음료(잔)", valAxisTitleFontSize: small ? 9 : 11, valAxisTitleColor: C.muted,
    showCatAxisTitle: true, catAxisTitle: "최고기온(℃)", catAxisTitleFontSize: small ? 9 : 11, catAxisTitleColor: C.muted,
    catAxisLabelColor: C.muted, valAxisLabelColor: C.muted, catAxisLabelFontSize: small ? 9 : 11, valAxisLabelFontSize: small ? 9 : 11,
    catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, valAxisTitleFontFace: FONT, catAxisTitleFontFace: FONT,
    valGridLine: { color: "E1E4E8", size: 0.5 }, catGridLine: { style: "none" },
    plotArea: { fill: { color: "FFFFFF" } },
  });
}
card(s, 0.89, 1.98, 6.9, 4.1, { fill: "FFFFFF", line: C.border });
scatter(s, 1.0, 2.05, 6.7, 3.95);
const vs = [["①", "D열과 G열 함께 선택 (Ctrl + 클릭)"], ["②", "삽입 > 차트 > 분산형"], ["③", "점 우클릭 > 추세선 추가 > 선형"], ["④", "✓ 수식을 차트에 표시 · ✓ R-제곱 값 표시"], ["⑤", "차트 제목과 축 제목(단위) 넣기"]];
vs.forEach(([n, t], i) => {
  const last = i === 4;
  card(s, 8.05, 1.98 + i * 0.74, 4.4, 0.6, { fill: last ? C.navy : C.card, line: last ? null : C.border, r: 0.1 });
  T(s, [run(n + "  ", { bold: true, color: last ? C.orange : C.teal }), run(t, { color: last ? C.cream : C.navy })], { x: 8.25, y: 1.98 + i * 0.74 + 0.16, w: 4.1, h: 0.3, fontSize: 13 });
});
T(s, "추세선 식 y = 15.035x − 209.59, R² = 0.7781 → 회귀분석 결과와 같습니다", { x: 8.05, y: 5.72, w: 4.4, h: 0.4, fontSize: 12, color: C.muted });
footer(s, FOOT_DATA);
s.addNotes("엑셀 추세선의 수식과 R²는 회귀분석 도구 결과와 완전히 같으므로, 서로 검산하는 용도로도 쓸 수 있습니다. 분산형 차트는 왼쪽 열이 X축이 되므로 D열(기온)이 G열보다 앞에 있어야 합니다. 축 제목에 단위(℃, 잔)를 꼭 넣게 해 주세요.");

// =====================================================================
// 12. 시각화 ② 콤보 + 차트 고르기
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "시각화 ② 추이와 비교");
title(s, "보고 목적에 맞는 차트를 고릅니다");
const days = data.map((d) => String(Number(d["날짜"].slice(8))));
card(s, 0.89, 2.06, 7.0, 4.0, { fill: "FFFFFF", line: C.border });
s.addChart([
  { type: pres.charts.BAR, data: [{ name: "아이스음료(잔)", labels: days, values: ice }], options: { chartColors: ["9FC3C8"], barGapWidthPct: 40 } },
  { type: pres.charts.LINE, data: [{ name: "최고기온(℃)", labels: days, values: temps }], options: { chartColors: [C.brown], lineSize: 2, lineDataSymbol: "none", secondaryValAxis: true, secondaryCatAxis: true } },
], {
  x: 1.0, y: 2.12, w: 6.8, h: 3.88, objectName: "DailyCombo",
  showTitle: true, title: "9월 일별 아이스음료 판매량과 최고기온", titleFontSize: 13, titleColor: C.navy, titleFontFace: FONT,
  showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: C.muted, legendFontFace: FONT,
  valAxes: [
    { showValAxisTitle: true, valAxisTitle: "잔", valAxisMinVal: 0, valAxisMaxVal: 300, valAxisLabelColor: C.muted, valAxisLabelFontSize: 10, valAxisTitleFontSize: 10, valAxisTitleColor: C.muted, valGridLine: { color: "E1E4E8", size: 0.5 } },
    { showValAxisTitle: true, valAxisTitle: "℃", valAxisMinVal: 15, valAxisMaxVal: 35, valAxisLabelColor: C.muted, valAxisLabelFontSize: 10, valAxisTitleFontSize: 10, valAxisTitleColor: C.muted, valGridLine: { style: "none" } },
  ],
  catAxes: [{ catAxisTitle: "일", showCatAxisTitle: false, catAxisLabelColor: C.muted, catAxisLabelFontSize: 9 }, { catAxisHidden: true }],
});
const guide = [["추이", "꺾은선", "날짜별 변화"], ["관계", "분산형 + 추세선", "두 변수의 연관"], ["비교", "막대", "평일 vs 주말 방문객"], ["여러 변수", "상관표 + 색조", "전체 관계 한눈에"]];
guide.forEach(([k, t, d], i) => {
  const x = 8.1 + (i % 2) * 2.24, y = 2.06 + Math.floor(i / 2) * 1.42;
  card(s, x, y, 2.11, 1.28);
  T(s, k, { x: x + 0.18, y: y + 0.14, w: 1.8, h: 0.28, fontSize: 13, bold: true, color: C.brown });
  T(s, t, { x: x + 0.18, y: y + 0.45, w: 1.85, h: 0.3, fontSize: 14, bold: true, color: C.navy });
  T(s, d, { x: x + 0.18, y: y + 0.8, w: 1.85, h: 0.4, fontSize: 12, color: C.body });
});
card(s, 8.1, 4.97, 4.35, 1.09, { fill: C.navy, line: null, r: 0.1 });
T(s, "콤보 차트 만들기", { x: 8.3, y: 5.08, w: 4.0, h: 0.28, fontSize: 13, bold: true, color: C.orange });
T(s, "차트 우클릭 > 차트 종류 변경 > 콤보 > 최고기온의 ‘보조 축’에 ✓", { x: 8.3, y: 5.4, w: 4.0, h: 0.6, fontSize: 12, color: C.cream });
footer(s, FOOT_DATA);
s.addNotes("단위가 다른 두 값(잔, ℃)을 한 차트에 넣을 때는 보조 축을 씁니다. 날씨가 서늘해지는 9월 말로 갈수록 막대가 낮아지는 흐름을 함께 읽어 보세요. 평일·주말 평균 방문객(239명/301명)은 AVERAGEIF(C2:C31,0,F2:F31)처럼 구해 막대 차트로 비교할 수 있습니다.");

// =====================================================================
// 13. AI 보고서 프롬프트
s = pres.addSlide(); s.background = { color: C.sky };
label(s, "종합 예제 · 보고서 ① 프롬프트");
title(s, "엑셀 결과 수치를 넣어 보고서 초안을 요청합니다", { size: 28 });
card(s, 0.89, 1.94, 11.56, 2.05, { fill: C.navy, line: null });
T(s, "요청 1 · 보고서", { x: 1.17, y: 2.13, w: 11, h: 0.28, fontSize: 13, bold: true, color: C.orange });
T(s, "너는 베이커리 카페 본사 운영지원팀 직원이야. 해운대점 9월 일별 데이터(30일)를 엑셀로 분석한 결과로 팀장님께 드릴 1쪽 분석 보고서를 써 줘. 상관계수: 최고기온–아이스음료 0.88, 강수량–방문객수 −0.64, 최고기온–방문객수 0.18. 단순회귀: 아이스음료 = 15.04×최고기온 − 209.6, R² 0.78, 28℃ 예측 211잔. 다중회귀(방문객수): 주말 +87.1명, 강수량 1mm당 −2.84명, 최고기온 1℃당 +2.65명, R² 0.93. 형식은 제목, 요약 3줄, 핵심 지표 표, 시사점 3가지, 분석의 한계 1줄. 주어진 숫자 외의 수치는 만들지 말고, 상관관계를 원인처럼 쓰지 마.",
  { x: 1.17, y: 2.46, w: 11.05, h: 1.55, fontSize: 12 , color: C.cream });
[[0.89, "요청 2 · 근거 표시", "각 문장 끝에 근거가 된 수치를 괄호로 표시해 줘. 확실하지 않은 문장은 [확인 필요]라고 붙여 줘."],
 [6.75, "요청 3 · 실행 제안", "시사점을 매장 운영(재료 준비, 직원 배치, 프로모션)에 바로 쓸 수 있는 행동 3가지로 바꿔 줘."]].forEach(([x, h, d]) => {
  card(s, x, 4.36, 5.69, 1.24, { fill: C.navy, line: null });
  T(s, h, { x: x + 0.25, y: 4.55, w: 5.3, h: 0.28, fontSize: 13, bold: true, color: C.orange });
  T(s, d, { x: x + 0.25, y: 4.88, w: 5.3, h: 0.6, fontSize: 12, color: C.cream });
});
T(s, "포인트 · 계산은 엑셀이, 문장은 AI가. 원본 30행 대신 분석 결과 숫자만 넣습니다 (7교시 보안).", { x: 0.89, y: 5.86, w: 11.9, h: 0.28, fontSize: 13, bold: true, color: C.navy });
footer(s);
s.addNotes("3교시의 5요소(역할·작업·맥락·형식·조건)가 모두 들어 있는지 학생들이 찾아보게 하세요. 특히 조건의 ‘주어진 숫자 외의 수치는 만들지 마’, ‘상관관계를 원인처럼 쓰지 마’가 다음 장 검증에서 걸러야 할 오류를 미리 막는 장치입니다. 요청 2의 근거 표시는 7교시 검증 프롬프트를 응용한 것입니다.");

// =====================================================================
// 14. 검증
s = pres.addSlide(); s.background = { color: C.cream };
label(s, "종합 예제 · 보고서 ② 검증");
title(s, "AI가 쓴 해석, 이렇게 고쳤습니다");
const ok = { text: "✓ 유지", options: { color: C.teal, bold: true, align: "center" } };
const bad = (t) => ({ text: "✗ " + t, options: { color: C.brown, bold: true, align: "center" } });
table(s, [
  ["AI가 쓴 문장", "판정", "이유 → 고친 문장"],
  ["기온과 아이스음료 판매는 강한 양의 상관관계(r=0.88)를 보인다.", ok, "엑셀 CORREL 결과와 일치"],
  ["기온 상승이 아이스음료 판매 증가의 원인임이 확인되었다.", bad("수정"), "상관 ≠ 인과 → “기온이 높은 날 판매가 많은 경향이 있다”"],
  ["R² 0.78이므로 예측 정확도는 78%이다.", bad("수정"), "R²는 정확도가 아님 → “판매량 변동의 78%를 기온으로 설명”"],
  ["폭염(40℃)에는 약 392잔이 팔릴 것으로 예상된다.", bad("삭제"), "관측 범위(23.5~31.5℃) 밖의 예측"],
  ["10월에는 매출이 10% 증가할 전망이다.", bad("삭제"), "데이터에 없는 수치 (환각)"],
], { x: 0.89, y: 2.06, w: 11.56, colW: [5.0, 1.15, 5.41], rowH: [0.42, 0.5, 0.5, 0.5, 0.5, 0.5], fontSize: 13 });
card(s, 0.89, 5.35, 11.56, 0.68, { fill: C.navy, line: null, r: 0.1 });
T(s, [run("검증 세 가지 · ", { color: C.cream }), run("숫자는 엑셀과 대조 · 범위 밖 예측 금지 · ‘원인’ 대신 ‘경향·관계’", { bold: true, color: C.orange })],
  { x: 1.12, y: 5.53, w: 11.1, h: 0.32, fontSize: 15 });
footer(s);
s.addNotes("실제 생성형 AI가 자주 만드는 오류 유형입니다. 392잔은 회귀식에 40℃를 넣으면 계산은 맞지만, 관측하지 않은 범위라 믿을 수 없다는 점이 핵심입니다. 7교시 검증 체크리스트(사실·수치·출처·날짜·맥락) 중 ‘수치’와 ‘맥락’ 항목에 해당한다고 연결해 주세요.");

// =====================================================================
// 15. 완성 보고서
s = pres.addSlide(); s.background = { color: C.sky };
label(s, "종합 예제 · 보고서 ③ 완성본", 0.78);
card(s, 0.89, 1.25, 11.55, 5.4, { fill: C.card, line: C.pale, r: 0.1 });
T(s, "해운대점 9월 판매 데이터 분석 보고", { x: 1.17, y: 1.45, w: 7, h: 0.33, fontSize: 18, bold: true, color: C.navy, objectName: "Title" });
T(s, "운영지원팀 · 2026.10.19.", { x: 9.93, y: 1.52, w: 2.23, h: 0.26, fontSize: 12, color: C.muted, align: "right" });
T(s, [run("요약", { bold: true, color: C.navy }), run("　① 아이스음료 판매는 최고기온과 강한 양의 상관(r=0.88)이 있고, 1℃ 오를 때 약 15잔 늘어나는 경향　② 방문객 수는 주말(+87명)과 강수량(1mm당 −2.8명)의 영향이 크고 기온 영향은 작음　③ 방문객 수와 매출은 거의 같이 움직임(r=0.98)", { color: C.body })],
  { x: 1.17, y: 1.88, w: 11.0, h: 0.6, fontSize: 12 });
scatter(s, 1.12, 2.58, 5.35, 2.45, { titleText: "최고기온–아이스음료 산점도와 추세선", small: true });
table(s, [
  ["핵심 지표", "값", "엑셀 근거"],
  ["상관계수 r (기온·아이스음료)", "0.88", "CORREL"],
  ["회귀식", "15.04x − 209.6", "SLOPE·INTERCEPT"],
  ["결정계수 R²", "0.78", "RSQ"],
  ["28℃ 예측", "211잔", "FORECAST.LINEAR"],
  ["방문객 다중회귀 R²", "0.93", "회귀 분석"],
], { x: 6.72, y: 2.6, w: 5.45, colW: [2.35, 1.45, 1.65], rowH: 0.38, fontSize: 11 });
T(s, [run("시사점", { bold: true, color: C.navy }), run("　① 다음 날 최고기온 예보로 아이스음료 재료 준비량 조정(28℃ ≈ 211잔)　② 주말은 평일보다 방문객이 평균 62명 많아(239→301명) 인력 보강 검토　③ 비 예보일은 방문객 감소에 대비해 포장·배달 프로모션 검토", { color: C.body })],
  { x: 1.17, y: 5.17, w: 11.0, h: 0.6, fontSize: 12 });
T(s, [run("한계", { bold: true, color: C.navy }), run("　9월 30일 자료(최고기온 23.5~31.5℃)로 분석해 다른 계절·지점에 그대로 적용할 수 없으며, 상관관계는 원인을 뜻하지 않음", { color: C.body })],
  { x: 1.17, y: 5.9, w: 11.0, h: 0.5, fontSize: 12 });
footer(s, FOOT_DATA);
s.addNotes("검증을 거친 최종본입니다. 모든 수치가 엑셀 결과(상관표·회귀 출력·추세선)와 일치하는지 표의 ‘엑셀 근거’ 열로 대조할 수 있게 했습니다. 8교시 예시 답안처럼 마지막에 ‘한계’를 스스로 밝히는 문장이 좋은 분석 보고서의 조건이라는 점을 강조하고, 이 보고서를 8교시처럼 메일·공지로 바꾸는 것을 과제로 내도 좋습니다. 평일 평균 239.2명, 주말 평균 300.9명.");

pres.writeFile({ fileName: OUT }).then(() => console.log("written", OUT));
