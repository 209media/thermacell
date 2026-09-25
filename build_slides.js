// 2 slajdy B2B PL do „1_CEE kick off FY27 ze sprzedażą” – styl prezentacji (tytuł Calibri Bold, turkus 108C96, granat 174769).
// Budowa: NODE_PATH=<node_modules> node build_slides.js
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

// ---- DANE: raport KPI B2B z 25.08.2026 (Sage ZORDDET, WEB-B2B, narastająco) ----
const KPI = { date: '25.08.2026', clients: 174, withOrder: 127, orders: 308, sales: 396757 };
// ---- Nowe zamówienia z panelu sbm-partners.com po 25.08 (do uzupełnienia, gdy będzie eksport) ----
// Zrzut listy zamówień z panelu, nr 323–331 (31.08–25.09.2026), wartości netto w PLN; nr 322 z 25.08 jest już w raporcie.
// 108 402,67 zł, po 2,5% rabacie za przelew 7 dni (nr 327, 328) 108 288,17 zł × 0,236967 EUR/PLN (kurs z prezentacji) = 25 661 €.
const PANEL = { from: '26.08', to: '25.09.2026', orders: 9, sales: 25661, newClientsWithOrder: 0 };
const BC = { clients: 166, orders: 928, sales: 464000, aov: 500, start: new Date('2026-03-16'), end: new Date('2026-10-01') };
const FY27 = { sales: '900 tys. €', orders: '1 125', aov: '800 €', newStores: '350', active: '95%' };

const T = {
  orders: KPI.orders + PANEL.orders,
  sales: KPI.sales + PANEL.sales,
  withOrder: KPI.withOrder + PANEL.newClientsWithOrder,
  clients: KPI.clients,
};
T.aov = Math.round(T.sales / T.orders);
T.noOrder = T.clients - T.withOrder;
const asOf = PANEL.orders ? PANEL.to : KPI.date;
const asOfDate = new Date(asOf.split('.').reverse().join('-'));
const periodPct = (asOfDate - BC.start) / (BC.end - BC.start);
const pct = x => Math.round(x * 100) + '%';
const vsPlan = () => T.sales >= BC.sales * periodPct ? 'przed planem do dziś' : `${pct(T.sales / (BC.sales * periodPct))} planu do dziś`;
const nbsp = n => String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
const eur = n => nbsp(n) + ' €';
const ktys = n => (n / 1000).toFixed(1).replace('.', ',') + ' tys. €';
const x100 = (a, b) => (a / b).toFixed(1).replace('.', ',') + '×';

// Kolory z prezentacji
const C = { teal: '108C96', navy: '174769', deep: '001F5F', ink: '30597A', tint: 'DAEBEA', line: '00676D', ok: '0E8B95', bad: 'C0392B', amber: 'F8AF2D', amberTxt: 'B7791F', white: 'FFFFFF', grey: 'E3ECEC', muted: '6B7F8E' };
const F = 'Calibri';

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: '#' + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}
const bullets = (items, extra = {}) => items.map((t, j) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: j < items.length - 1, paraSpaceAfter: 5, ...extra } }));
const title = (s, text) => s.addText(text, { x: 0.45, y: 0.3, w: 12.4, h: 0.8, fontFace: F, fontSize: 40, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5, jak prezentacja
  pres.title = 'B2B PL – kick off FY27';

  // ================= SLAJD 1: założenia, realizacja, co zadziałało / nie, learningi =================
  const s1 = pres.addSlide();
  s1.background = { color: C.white };
  title(s1, 'B2B PL – PODSUMOWANIE FY26');

  const cols = [
    { h: 'Założenia', ic: fa.FaBullseye, items: [
      'Własny kanał zamówień dla sklepów i hurtu: panel sbm-partners.com, 1.\u00A0zamówienie 19.03.2026',
      'Business case FY26: 166 klientów, 464 tys. €, 928 zamówień, średnio 500 €',
    ] },
    { h: 'Jak to zrobiliśmy', ic: fa.FaCogs, items: [
      'Cała baza na panelu: rejestracja, konta, grupy cenowe',
      'Ceny netto z rabatem klienta; płatność 90 dni lub 7 dni z 2,5% rabatu',
      'Od 12.06 klienci przypisani do KAM + cotygodniowy raport KPI',
    ] },
    { h: 'Co zadziałało', ic: fa.FaCheck, items: [
      `${T.clients} kont (${pct(T.clients / BC.clients)} bazy), ${T.withOrder} z zamówieniem`,
      `${ktys(T.sales)} = ${pct(T.sales / BC.sales)} budżetu, ${pct(T.sales / (BC.sales * periodPct))} planu do dziś`,
      '59% klientów wraca z 2.\u00A0zamówieniem w ciągu 60\u00A0dni',
      'Część klientów zamawia sama, bez telefonu do PH',
    ] },
    { h: 'Co nie zadziałało', ic: fa.FaTimes, items: [
      `${T.orders} zamówień = ${pct(T.orders / BC.orders)} celu: rzadkie, duże zamówienia „na stock”`,
      'Brak kampanii B2B i kontaktu w sezonie (newsletter, SMS)',
      'Korekty cen ręcznie w Sage: klient widzi inne kwoty niż na fakturze, brak śledzenia przesyłek',
      'Brak programu lojalnościowego, niewygodny panel',
    ] },
  ];
  const cx0 = 0.45, cw = 2.95, cg = 0.2, hy = 1.3, hh = 0.5, by = 1.9, bh = 3.15;
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i], x = cx0 + i * (cw + cg);
    s1.addShape(pres.shapes.RECTANGLE, { x, y: hy, w: cw, h: hh, fill: { color: C.navy }, line: { color: C.navy } });
    s1.addImage({ data: await icon(c.ic, C.white), x: x + 0.18, y: hy + 0.13, w: 0.24, h: 0.24 });
    s1.addText(c.h, { x: x + 0.52, y: hy, w: cw - 0.6, h: hh, fontFace: F, fontSize: 16, bold: true, color: C.white, valign: 'middle', margin: 0, isTextBox: true });
    s1.addShape(pres.shapes.RECTANGLE, { x, y: by, w: cw, h: bh, fill: { color: C.white }, line: { color: C.navy, width: 1.25 } });
    s1.addText(bullets(c.items), { x: x + 0.15, y: by + 0.15, w: cw - 0.28, h: bh - 0.25, fontFace: F, fontSize: 12.5, color: C.ink, valign: 'top', margin: 0, isTextBox: true });
  }

  // Learningi FY27
  const ly = 5.25, lh = 1.85;
  s1.addShape(pres.shapes.RECTANGLE, { x: cx0, y: ly, w: 12.4, h: lh, fill: { color: C.navy }, line: { color: C.navy } });
  s1.addText('LEARNINGI NA FY27', { x: cx0 + 0.3, y: ly + 0.15, w: 6, h: 0.35, fontFace: F, fontSize: 16, bold: true, color: C.amber, margin: 0, isTextBox: true });
  const learn = [
    'Kampania B2B i budżet mediowy, nie tylko pozyskanie przez PH',
    'Program partnerski: punkty za zgody marketingowe i częstsze zamówienia',
    'Prostszy panel + codzienny import z Sage: kwoty jak na fakturze, nr listu przewozowego',
    'Liczymy konta, które zamawiają, i częstotliwość, nie rejestracje',
  ];
  const lw = 2.9;
  for (let i = 0; i < learn.length; i++) {
    const x = cx0 + 0.3 + i * (lw + 0.12);
    s1.addShape(pres.shapes.OVAL, { x, y: ly + 0.68, w: 0.42, h: 0.42, fill: { color: C.amber }, line: { color: C.amber } });
    s1.addText(String(i + 1), { x, y: ly + 0.68, w: 0.42, h: 0.42, fontFace: F, fontSize: 15, bold: true, color: C.navy, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    s1.addText(learn[i], { x: x + 0.52, y: ly + 0.58, w: lw - 0.58, h: 1.15, fontFace: F, fontSize: 12.5, color: C.white, valign: 'top', margin: 0, isTextBox: true });
  }
  s1.addNotes(`Jeden slajd zamiast angielskich slajdów o panelu B2B. Założenia z business case FY26: 166 klientów, 928 zamówień, 464 tys. €. Panel się przyjął: ${T.clients} kont, ${pct(T.sales / BC.sales)} budżetu sprzedaży (${vsPlan()}), 59% klientów wraca z drugim zamówieniem. Słabo z liczbą zamówień (${T.orders} = ${pct(T.orders / BC.orders)} celu): klienci zamawiają rzadko i dużo, bez kampanii, bez kontaktu w sezonie, a korekty cen w Sage nie wracały do panelu. Stąd 4 learningi na FY27.`);

  // ================= SLAJD 2: wynik + start nowej platformy =================
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  title(s2, 'B2B PL – WYNIK I START FY27');
  s2.addText(`Wynik FY26 vs budżet · stan na ${asOf}`, { x: 0.45, y: 1.25, w: 7, h: 0.4, fontFace: F, fontSize: 16, bold: true, color: C.navy, margin: 0, isTextBox: true });

  const kpis = [
    { v: eur(T.sales), l: 'Sprzedaż netto', b: `budżet ${eur(BC.sales)} · ${pct(T.sales / (BC.sales * periodPct))} planu do dziś`, p: T.sales / BC.sales, pl: pct(T.sales / BC.sales), col: C.ok },
    { v: `${T.withOrder} / ${T.clients}`, l: 'Klienci z zamówieniem', b: `${T.noOrder} kont bez zamówienia`, p: T.withOrder / T.clients, pl: pct(T.withOrder / T.clients), col: C.ok },
    { v: nbsp(T.orders), l: 'Zamówienia', b: `budżet ${nbsp(BC.orders)}`, p: T.orders / BC.orders, pl: pct(T.orders / BC.orders), col: C.bad },
    { v: eur(T.aov), l: 'Średnia wartość zamówienia', b: `plan ${eur(BC.aov)} (${x100(T.aov, BC.aov)})`, p: 1.0, pl: pct(T.aov / BC.aov), col: C.amber },
  ];
  const ky0 = 1.8, kh = 1.0, bx = 3.5, bw = 3.0;
  kpis.forEach((k, i) => {
    const y = ky0 + i * kh;
    s2.addText(k.v, { x: 0.45, y, w: 2.95, h: 0.55, fontFace: F, fontSize: 28, bold: true, color: C.navy, margin: 0, valign: 'bottom', isTextBox: true });
    s2.addText(k.l, { x: 0.45, y: y + 0.56, w: 2.95, h: 0.3, fontFace: F, fontSize: 12, color: C.muted, margin: 0, isTextBox: true });
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.28, w: bw, h: 0.26, fill: { color: C.grey }, line: { color: C.grey }, rectRadius: 0.13 });
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.28, w: Math.max(0.26, bw * k.p), h: 0.26, fill: { color: k.col }, line: { color: k.col }, rectRadius: 0.13 });
    s2.addText(k.pl, { x: bx + bw + 0.1, y: y + 0.2, w: 0.8, h: 0.42, fontFace: F, fontSize: 16, bold: true, color: k.col === C.amber ? C.amberTxt : k.col, margin: 0, valign: 'middle', isTextBox: true });
    s2.addText(k.b, { x: bx, y: y + 0.58, w: bw + 0.9, h: 0.28, fontFace: F, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
  });

  const fy = 5.95;
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: fy, w: 6.95, h: 0.85, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s2.addText([
    { text: 'Cele FY27:  ', options: { bold: true, color: C.navy } },
    { text: `${FY27.sales} · ${FY27.orders} zamówień · średnio ${FY27.aov} · ${FY27.newStores} nowych sklepów · ${FY27.active} bazy aktywnej`, options: { color: C.ink } },
  ], { x: 0.65, y: fy, w: 6.6, h: 0.85, fontFace: F, fontSize: 13, valign: 'middle', margin: 0, isTextBox: true });

  // Nowa platforma
  const px = 7.75, pw = 5.1, py = 1.25, ph = 5.55;
  s2.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: C.navy }, line: { color: C.navy } });
  s2.addText('NOWA PLATFORMA B2B', { x: px + 0.3, y: py + 0.22, w: pw - 0.6, h: 0.4, fontFace: F, fontSize: 18, bold: true, color: C.amber, margin: 0, isTextBox: true });
  s2.addText('start 1.10.2026 · stary panel sbm-partners.com działa równolegle do 30.10', { x: px + 0.3, y: py + 0.62, w: pw - 0.6, h: 0.5, fontFace: F, fontSize: 12, italic: true, color: 'C9DCE8', margin: 0, isTextBox: true });
  const feats = [
    ['Ceny netto z rabatem klienta', 'kwota w koszyku = kwota na fakturze'],
    ['Kredyt kupiecki z limitem', 'termin płatności w koszyku: krótszy = taniej'],
    ['Zamówienie w 4 krokach', 'katalog, ilość, płatność, data; także na telefonie'],
    ['Automatyczne maile do klienta', 'aktywacja konta, 1. zamówienie, potwierdzenia'],
    ['Program partnerski od 1.10', 'wersja ręczna, silnik i integracja w Q4'],
  ];
  const fx = px + 0.3, fy0 = py + 1.3, fstep = 0.82;
  const chk = await icon(fa.FaCheckCircle, C.amber);
  feats.forEach((f, i) => {
    const y = fy0 + i * fstep;
    s2.addImage({ data: chk, x: fx, y: y + 0.03, w: 0.26, h: 0.26 });
    s2.addText([
      { text: f[0], options: { bold: true, color: C.white, breakLine: true } },
      { text: f[1], options: { color: 'D5E3EC', fontSize: 12 } },
    ], { x: fx + 0.4, y, w: pw - 0.9, h: 0.75, fontFace: F, fontSize: 14, valign: 'top', margin: 0, isTextBox: true });
  });

  s2.addText(`Dane: raport KPI B2B z ${KPI.date} (Sage ZORDDET, WEB-B2B, narastająco)${PANEL.orders ? ` + nowe zamówienia z panelu sbm-partners.com ${PANEL.from}–${PANEL.to}` : ''}. 59% powrotów: Sage ZORDDET 16.03–22.07.2026.`, { x: 0.45, y: 6.95, w: 12.4, h: 0.3, fontFace: F, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
  s2.addNotes(`Wynik na ${asOf}: ${pct(T.sales / BC.sales)} budżetu sprzedaży (${pct(T.sales / (BC.sales * periodPct))} planu do dziś), ${pct(T.withOrder / T.clients)} kont z zamówieniem, ale tylko ${pct(T.orders / BC.orders)} celu zamówień. Średnie zamówienie ${x100(T.aov, BC.aov)} plan, bo klienci zamawiają rzadko i dużo. 1 października startuje nowa platforma: ceny i rabaty od razu w koszyku, kredyt kupiecki, rabat za krótszy termin płatności, automatyczne maile i program partnerski. Cele FY27: ${FY27.sales}, ${FY27.orders} zamówień.`);

  await pres.writeFile({ fileName: 'B2B_PL_kickoff_FY27_2_slajdy.pptx' });
  console.log('ok');
})();
