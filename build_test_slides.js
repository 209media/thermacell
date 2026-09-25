// 2 slajdy: podsumowanie testu panelu B2B FY26 (co zagrało / co nie) – styl „1_CEE kick off FY27 ze sprzedażą”.
// Budowa: NODE_PATH=<node_modules> node build_test_slides.js
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
// Przyrost od przypisania klientów do KAM (12.06), zakładka „Śledzenie” raportu KPI
const KAM = { orders: 28 + 55, sales: 26077 + 35131 };

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

// Kolory z prezentacji
const C = { teal: '108C96', navy: '174769', ink: '30597A', tint: 'DAEBEA', line: '00676D', ok: '0E8B95', bad: 'C0392B', badTint: 'FBEDEB', amber: 'F8AF2D', white: 'FFFFFF', muted: '6B7F8E', soft: 'F3F7F7' };
const F = 'Calibri';

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: '#' + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}
const bullets = (items, extra = {}) => items.map((t, j) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: j < items.length - 1, paraSpaceAfter: 6, ...extra } }));
const title = (s, text) => s.addText(text, { x: 0.45, y: 0.3, w: 12.4, h: 0.75, fontFace: F, fontSize: 36, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });
const subtitle = (s, text) => s.addText(text, { x: 0.45, y: 1.02, w: 12.4, h: 0.32, fontFace: F, fontSize: 14, color: C.muted, margin: 0, isTextBox: true });
const source = (s, text) => s.addText(text, { x: 0.45, y: 7.05, w: 12.4, h: 0.28, fontFace: F, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5, jak prezentacja
  pres.title = 'Test panelu B2B FY26';

  // ================= SLAJD 1: CO ZAGRAŁO =================
  const s1 = pres.addSlide();
  s1.background = { color: C.white };
  title(s1, 'TEST PANELU B2B FY26 – CO ZAGRAŁO');
  subtitle(s1, `Panel sbm-partners.com, pierwsze zamówienie 19.03.2026 · wyniki narastająco na ${asOf}`);

  const tiles = [
    { v: nbsp(T.clients), l: `kont na panelu = ${pct(T.clients / BC.clients)} założonej bazy (${BC.clients})` },
    { v: `${T.withOrder}`, l: `klientów z zamówieniem = ${pct(T.withOrder / T.clients)} kont` },
    { v: ktys(T.sales), l: `sprzedaży netto = ${pct(T.sales / BC.sales)} budżetu, ${pct(T.sales / (BC.sales * periodPct))} planu do dziś` },
    { v: '59%', l: 'klientów wraca z 2. zamówieniem w ciągu 60 dni' },
  ];
  const tw = 2.95, tg = 0.2, ty = 1.55, th = 1.3;
  tiles.forEach((t, i) => {
    const x = 0.45 + i * (tw + tg);
    s1.addShape(pres.shapes.RECTANGLE, { x, y: ty, w: tw, h: th, fill: { color: C.white }, line: { color: C.teal, width: 1.5 } });
    s1.addText(t.v, { x: x + 0.2, y: ty + 0.12, w: tw - 0.4, h: 0.6, fontFace: F, fontSize: 30, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });
    s1.addText(t.l, { x: x + 0.2, y: ty + 0.72, w: tw - 0.35, h: 0.5, fontFace: F, fontSize: 12, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
  });

  const good = [
    { h: 'Klienci przyjęli panel', ic: fa.FaUsers, items: [
      `Cała baza zarejestrowana, ${T.clients} kont`,
      'Część klientów zamawia sama, bez telefonu do PH',
      'Klienci wracają: 59% składa 2. zamówienie w ciągu 60 dni',
    ] },
    { h: 'Panel i warunki działały', ic: fa.FaDesktop, items: [
      'Rejestracja, konta i grupy cenowe bez problemów',
      'Ceny netto z indywidualnym rabatem w katalogu',
      'Płatność 90 dni albo 7 dni z 2,5% rabatu',
    ] },
    { h: 'KAM i kontrola wyniku', ic: fa.FaChartLine, items: [
      `Od 12.06 klienci przypisani do KAM: +${KAM.orders} zamówień i ${ktys(KAM.sales)}`,
      'Cotygodniowy raport KPI: wynik na tle BC co tydzień',
      `${pct(T.sales / BC.sales)} budżetu rocznego na tydzień przed końcem BC`,
    ] },
  ];
  const gw = 4.0, gg = 0.2, gy = 3.1, gh = 2.55, hh = 0.55;
  for (let i = 0; i < good.length; i++) {
    const c = good[i], x = 0.45 + i * (gw + gg);
    s1.addShape(pres.shapes.RECTANGLE, { x, y: gy, w: gw, h: hh, fill: { color: C.navy }, line: { color: C.navy } });
    s1.addImage({ data: await icon(c.ic, C.white), x: x + 0.2, y: gy + 0.14, w: 0.27, h: 0.27 });
    s1.addText(c.h, { x: x + 0.6, y: gy, w: gw - 0.7, h: hh, fontFace: F, fontSize: 17, bold: true, color: C.white, valign: 'middle', margin: 0, isTextBox: true });
    s1.addShape(pres.shapes.RECTANGLE, { x, y: gy + hh, w: gw, h: gh - hh, fill: { color: C.white }, line: { color: C.navy, width: 1.25 } });
    s1.addText(bullets(c.items), { x: x + 0.2, y: gy + hh + 0.2, w: gw - 0.4, h: gh - hh - 0.3, fontFace: F, fontSize: 14, color: C.ink, valign: 'top', margin: 0, isTextBox: true });
  }
  // Wniosek
  const wy1 = 5.85, wh1 = 1.0;
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: wy1, w: 12.4, h: wh1, fill: { color: C.navy }, line: { color: C.navy } });
  s1.addText([
    { text: 'WNIOSEK:  ', options: { bold: true, color: C.amber } },
    { text: 'panel jako kanał się sprawdził, klienci go używają i wracają. W FY27 skalujemy: nowa platforma od 1.10, program partnerski i kampania B2B.', options: { color: C.white } },
  ], { x: 0.75, y: wy1, w: 11.9, h: wh1, fontFace: F, fontSize: 15, valign: 'middle', margin: 0, isTextBox: true });
  source(s1, `Dane: raport KPI B2B z ${KPI.date} (Sage ZORDDET, WEB-B2B)${PANEL.orders ? ` + nowe zamówienia z panelu ${PANEL.from}–${PANEL.to}` : ''}. Powroty: Sage ZORDDET 16.03–22.07.2026. Przyrost KAM: zakładka „Śledzenie”.`);
  s1.addNotes(`Test panelu trwał od 19 marca. Klienci przyjęli kanał: ${T.clients} kont, ${T.withOrder} z zamówieniem, ${ktys(T.sales)} sprzedaży, ${pct(T.sales / BC.sales)} budżetu rocznego (${vsPlan()}). 59% klientów wraca z drugim zamówieniem w ciągu 60 dni. Technicznie panel działał: rejestracja, rabaty, ceny netto, terminy płatności. Od przypisania klientów do KAM 12 czerwca przybyło ${KAM.orders} zamówień i ${ktys(KAM.sales)}.`);

  // ================= SLAJD 2: CO NIE ZAGRAŁO =================
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  title(s2, 'TEST PANELU B2B FY26 – CO NIE ZAGRAŁO');
  subtitle(s2, `${nbsp(T.orders)} zamówień = ${pct(T.orders / BC.orders)} celu (${nbsp(BC.orders)}): klienci zamawiają rzadko, dużo i na żądanie, średnio ${eur(T.aov)} zamiast ${eur(BC.aov)}`);

  const bad = [
    { h: 'Brak procesów sprzedażowych', items: [
      'Brak ustalonej pracy KAM z panelem: kiedy dzwonić, jak aktywować, jak podnosić częstotliwość',
      'KAM bez narzędzi analitycznych i materiałów do rozmowy z klientem',
      `${T.noOrder} kont bez zamówienia, nikt ich nie reaktywował`,
    ] },
    { h: 'Brak procesów pozyskania nowych klientów', items: [
      'Brak kampanii B2B i budżetu mediowego, nowe konta tylko z portfeli PH',
      'Brak bazy zgód marketingowych: nie było zachęty, żeby je zbierać',
      `Baza urosła tylko o ${T.clients - BC.clients} kont ponad BC (${BC.clients} → ${T.clients})`,
    ] },
    { h: 'Źle przygotowana polityka cenowa', items: [
      'Większość zamówień to deale spoza cennika w panelu',
      'Korekty cen ręcznie w Sage: duże obciążenie obsługi',
      'Korekty nie wracały do panelu: klient widział inne kwoty niż na fakturze',
    ] },
  ];
  const bw = 4.0, bg = 0.2, by = 1.55, bh = 3.0;
  bad.forEach((c, i) => {
    const x = 0.45 + i * (bw + bg);
    s2.addShape(pres.shapes.RECTANGLE, { x, y: by, w: bw, h: bh, fill: { color: C.white }, line: { color: C.bad, width: 1.5 } });
    s2.addShape(pres.shapes.OVAL, { x: x + 0.2, y: by + 0.2, w: 0.46, h: 0.46, fill: { color: C.bad }, line: { color: C.bad } });
    s2.addText(String(i + 1), { x: x + 0.2, y: by + 0.2, w: 0.46, h: 0.46, fontFace: F, fontSize: 16, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    s2.addText(c.h, { x: x + 0.8, y: by + 0.12, w: bw - 0.95, h: 0.62, fontFace: F, fontSize: 16, bold: true, color: C.bad, valign: 'middle', margin: 0, isTextBox: true });
    s2.addText(bullets(c.items), { x: x + 0.2, y: by + 0.9, w: bw - 0.4, h: bh - 1.0, fontFace: F, fontSize: 13, color: C.ink, valign: 'top', margin: 0, isTextBox: true });
  });

  // Dodatkowo
  const ay = 4.72, ah = 0.72;
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: ay, w: 12.4, h: ah, fill: { color: C.badTint }, line: { color: C.badTint } });
  s2.addText([
    { text: 'Dodatkowo:  ', options: { bold: true, color: C.bad } },
    { text: 'brak kontaktu z klientem w sezonie (newsletter, SMS)  ·  brak programu lojalnościowego  ·  niewygodny panel  ·  brak śledzenia przesyłek  ·  brak wyników KPI na żywo w panelu', options: { color: C.ink } },
  ], { x: 0.65, y: ay, w: 12.0, h: ah, fontFace: F, fontSize: 13, valign: 'middle', margin: 0, isTextBox: true });

  // Wnioski FY27
  const wy = 5.6, wh = 1.4;
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: wy, w: 12.4, h: wh, fill: { color: C.navy }, line: { color: C.navy } });
  s2.addText('WNIOSKI NA FY27', { x: 0.75, y: wy + 0.12, w: 5, h: 0.32, fontFace: F, fontSize: 15, bold: true, color: C.amber, margin: 0, isTextBox: true });
  const concl = [
    'Proces sprzedaży B2B: rola KAM w panelu, cele aktywacji i częstotliwości',
    'Proces pozyskania: kampania B2B, cold mailing, budżet mediowy, zgody',
    'Jeden cennik w panelu: deale jako rabaty w koszyku, import z Sage codziennie',
    'Program partnerski i stały kontakt z klientem w sezonie',
  ];
  const cw = 2.95;
  concl.forEach((t, i) => {
    const x = 0.75 + i * (cw + 0.05);
    s2.addShape(pres.shapes.OVAL, { x, y: wy + 0.56, w: 0.38, h: 0.38, fill: { color: C.amber }, line: { color: C.amber } });
    s2.addText(String(i + 1), { x, y: wy + 0.56, w: 0.38, h: 0.38, fontFace: F, fontSize: 14, bold: true, color: C.navy, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    s2.addText(t, { x: x + 0.48, y: wy + 0.5, w: cw - 0.55, h: 0.82, fontFace: F, fontSize: 12.5, color: C.white, valign: 'top', margin: 0, isTextBox: true });
  });
  source(s2, `Dane: raport KPI B2B z ${KPI.date} (Sage ZORDDET, WEB-B2B)${PANEL.orders ? ` + nowe zamówienia z panelu ${PANEL.from}–${PANEL.to}` : ''}; podsumowanie testu panelu B2B, lipiec 2026.`);
  s2.addNotes(`Liczba zamówień to główna luka: ${T.orders} wobec ${BC.orders} w planie. Trzy przyczyny: po pierwsze nie było procesu sprzedaży wokół panelu, KAM nie mieli ustalonego sposobu pracy ani narzędzi. Po drugie nie było procesu pozyskania nowych klientów, bez kampanii i budżetu baza rosła tylko z portfeli PH. Po trzecie polityka cenowa nie była gotowa: większość zamówień to deale spoza cennika, korygowane ręcznie w Sage, a klient widział w panelu inne kwoty niż na fakturze. Na FY27: proces sprzedaży, proces pozyskania, jeden cennik w panelu i program partnerski.`);

  await pres.writeFile({ fileName: 'B2B_PL_test_panelu_FY26_2_slajdy.pptx' });
  console.log('ok');
})();
