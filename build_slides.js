const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

const C = { dark: '1F4E3D', green: '2E7D5B', tint: 'EEF4F0', ink: '1E2A24', muted: '5E6B64', ok: '2E7D5B', bad: 'C0392B', amber: 'E0A100', white: 'FFFFFF', grey: 'D9E1DC' };
const FONT = 'Calibri', HFONT = 'Cambria';

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: '#' + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
  pres.title = 'B2B PL – kick off FY27';

  // ---------- SLAJD 1 ----------
  const s1 = pres.addSlide();
  s1.background = { color: C.white };
  s1.addText('B2B PL: założenia, realizacja, learningi na FY27', { x: 0.5, y: 0.35, w: 12.3, h: 0.75, fontFace: HFONT, fontSize: 28, bold: true, color: C.dark, margin: 0, isTextBox: true });

  const cards = [
    { t: 'Założenia FY26', ic: fa.FaBullseye, col: C.green, items: ['166 klientów na platformie (66 obecnych + 100 nowych)', '464 tys. € sprzedaży netto', '928 zamówień, średnio 500 € na zamówienie'] },
    { t: 'Jak to zrobiliśmy', ic: fa.FaCogs, col: C.green, items: ['Cała baza klientów przeniesiona na platformę', 'Onboarding przez przedstawicieli: cel 5 nowych kont tygodniowo na osobę', 'Cotygodniowy raport KPI (Power BI + tracker klientów)'] },
    { t: 'Co zadziałało', ic: fa.FaCheck, col: C.ok, items: ['Cała baza na platformie: 166 kont (100%)', '311,6\u00A0tys.\u00A0€ sprzedaży = 67% budżetu', 'Średnie zamówienie 1\u00A0574\u00A0€ (ponad 3× plan)'] },
    { t: 'Co nie zadziałało', ic: fa.FaTimes, col: C.bad, items: ['198 zamówień = 21% celu: rzadkie, duże zamówienia „na stock”', '55 kont bez żadnego zamówienia', 'Brak komunikacji po rejestracji i działań retencyjnych'] },
  ];
  const cw = 2.895, cg = 0.25, cy = 1.35, ch = 3.45;
  for (let i = 0; i < cards.length; i++) {
    const c = cards[i], x = 0.5 + i * (cw + cg);
    s1.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, fill: { color: C.tint }, line: { color: C.tint }, rectRadius: 0.08 });
    s1.addShape(pres.shapes.OVAL, { x: x + 0.25, y: cy + 0.25, w: 0.55, h: 0.55, fill: { color: C.white }, line: { color: c.col, width: 1.5 } });
    s1.addImage({ data: await icon(c.ic, c.col), x: x + 0.385, y: cy + 0.385, w: 0.28, h: 0.28 });
    s1.addText(c.t, { x: x + 0.95, y: cy + 0.25, w: cw - 1.1, h: 0.55, fontFace: FONT, fontSize: 16, bold: true, color: c.col === C.green ? C.dark : c.col, valign: 'middle', margin: 0, isTextBox: true });
    s1.addText(c.items.map((t, j) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < c.items.length - 1, paraSpaceAfter: 8 } })),
      { x: x + 0.2, y: cy + 1.0, w: cw - 0.35, h: ch - 1.15, fontFace: FONT, fontSize: 13, color: C.ink, valign: 'top', margin: 0, isTextBox: true });
  }

  // Learningi FY27
  const ly = 5.05, lh = 2.0;
  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: ly, w: 12.33, h: lh, fill: { color: C.dark }, line: { color: C.dark }, rectRadius: 0.08 });
  s1.addText('Learningi na FY27', { x: 0.8, y: ly + 0.18, w: 6, h: 0.4, fontFace: HFONT, fontSize: 18, bold: true, color: C.white, margin: 0, isTextBox: true });
  const learn = [
    'Rejestracja to start, nie meta: liczymy konta, które zamawiają',
    'Automatyczna komunikacja od 1. dnia + reaktywacja uśpionych kont',
    'Mniejsze, częstsze zamówienia: program lojalnościowy, rabat za szybszą płatność',
    'Nowe sklepy także z lead genu (cold mailing), nie tylko z bazy przedstawicieli',
  ];
  const lw = 2.85;
  for (let i = 0; i < learn.length; i++) {
    const x = 0.8 + i * (lw + 0.18);
    s1.addShape(pres.shapes.OVAL, { x, y: ly + 0.78, w: 0.46, h: 0.46, fill: { color: C.amber }, line: { color: C.amber } });
    s1.addText(String(i + 1), { x, y: ly + 0.78, w: 0.46, h: 0.46, fontFace: FONT, fontSize: 16, bold: true, color: C.dark, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    s1.addText(learn[i], { x: x + 0.58, y: ly + 0.66, w: lw - 0.62, h: 1.1, fontFace: FONT, fontSize: 13, color: C.white, valign: 'top', margin: 0, isTextBox: true });
  }
  s1.addNotes('Jeden slajd zamiast 9–14. Plan zakładał 166 kont i 928 zamówień. Platforma przyjęła się (cała baza jest na niej, 2/3 budżetu sprzedaży), ale klienci zamawiają rzadko i dużo. Wnioski na FY27: aktywacja zamiast samej rejestracji, automatyczna komunikacja, częstotliwość zamówień, nowe sklepy z lead genu.');

  // ---------- SLAJD 2 ----------
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  s2.addText('B2B PL: wynik i nowy panel B2B', { x: 0.5, y: 0.35, w: 12.3, h: 0.75, fontFace: HFONT, fontSize: 28, bold: true, color: C.dark, margin: 0, isTextBox: true });
  s2.addText('Wynik FY26 YTD vs budżet', { x: 0.5, y: 1.25, w: 6.8, h: 0.4, fontFace: FONT, fontSize: 16, bold: true, color: C.muted, margin: 0, isTextBox: true });

  const kpis = [
    { v: '311 557 €', l: 'Sprzedaż netto', b: 'budżet 464 000 €', p: 0.671, pl: '67%', col: C.green },
    { v: '111 / 166', l: 'Klienci z zamówieniem', b: '55 kont uśpionych', p: 0.669, pl: '67%', col: C.green },
    { v: '198', l: 'Zamówienia', b: 'budżet 928', p: 0.213, pl: '21%', col: C.bad },
    { v: '1 574 €', l: 'Średnia wartość zamówienia', b: 'plan 500 €', p: 1.0, pl: '315%', col: C.amber },
  ];
  const ky0 = 1.8, kh = 1.0, bx = 3.55, bw = 3.0;
  kpis.forEach((k, i) => {
    const y = ky0 + i * kh;
    s2.addText(k.v, { x: 0.5, y, w: 2.9, h: 0.55, fontFace: FONT, fontSize: 28, bold: true, color: C.ink, margin: 0, valign: 'bottom', isTextBox: true });
    s2.addText(k.l, { x: 0.5, y: y + 0.56, w: 2.9, h: 0.3, fontFace: FONT, fontSize: 12, color: C.muted, margin: 0, isTextBox: true });
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.28, w: bw, h: 0.26, fill: { color: C.grey }, line: { color: C.grey }, rectRadius: 0.13 });
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.28, w: Math.max(0.26, bw * k.p), h: 0.26, fill: { color: k.col }, line: { color: k.col }, rectRadius: 0.13 });
    s2.addText(k.pl, { x: bx + bw + 0.1, y: y + 0.2, w: 0.75, h: 0.42, fontFace: FONT, fontSize: 16, bold: true, color: k.col === C.amber ? 'A87800' : k.col, margin: 0, valign: 'middle', isTextBox: true });
    s2.addText(k.b, { x: bx, y: y + 0.58, w: bw, h: 0.28, fontFace: FONT, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
  });

  // Cele FY27
  const fy = 5.95;
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: fy, w: 6.85, h: 0.85, fill: { color: C.tint }, line: { color: C.tint }, rectRadius: 0.08 });
  s2.addText([
    { text: 'Cele FY27:  ', options: { bold: true, color: C.dark } },
    { text: '900\u00A0tys.\u00A0€ · 1\u00A0125 zamówień · średnio 800\u00A0€ · 350 nowych sklepów · 95% bazy aktywnej', options: { color: C.ink } },
  ], { x: 0.7, y: fy, w: 6.5, h: 0.85, fontFace: FONT, fontSize: 13, valign: 'middle', margin: 0, isTextBox: true });

  // Nowy panel
  const px = 7.75, pw = 5.08, py = 1.25, ph = 5.55;
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: C.dark }, line: { color: C.dark }, rectRadius: 0.08 });
  s2.addText('Nowy panel B2B', { x: px + 0.3, y: py + 0.25, w: pw - 0.6, h: 0.45, fontFace: HFONT, fontSize: 18, bold: true, color: C.white, margin: 0, isTextBox: true });
  s2.addText('sbm-sales.site · odpowiada na learningi z FY26', { x: px + 0.3, y: py + 0.7, w: pw - 0.6, h: 0.3, fontFace: FONT, fontSize: 12, italic: true, color: 'BFD8CC', margin: 0, isTextBox: true });
  const feats = [
    ['Ceny netto już z rabatem klienta', 'to, co w koszyku, jest na fakturze'],
    ['Kredyt kupiecki z limitem', 'zamiast przedpłaty, np. 14 dni'],
    ['Termin płatności w koszyku', 'krótszy termin = niższa cena od razu'],
    ['Zamówienie w 4 krokach', 'katalog, ilość, płatność, data; działa też na telefonie'],
    ['Automatyczne maile do klienta', 'aktywacja konta, instrukcja 1. zamówienia, potwierdzenia'],
  ];
  const fx = px + 0.3, fy0 = py + 1.2, fstep = 0.84;
  const chk = await icon(fa.FaCheckCircle, C.amber);
  feats.forEach((f, i) => {
    const y = fy0 + i * fstep;
    s2.addImage({ data: chk, x: fx, y: y + 0.03, w: 0.28, h: 0.28 });
    s2.addText([
      { text: f[0], options: { bold: true, color: C.white, breakLine: true } },
      { text: f[1], options: { color: 'D5E6DD', fontSize: 12 } },
    ], { x: fx + 0.42, y, w: pw - 0.95, h: 0.75, fontFace: FONT, fontSize: 14, valign: 'top', margin: 0, isTextBox: true });
  });

  s2.addText('Dane: raport KPI B2B, FY26 YTD (Power BI, WEB-B2B). Panel: funkcje na podstawie testów, wrzesień 2026.', { x: 0.5, y: 6.95, w: 12.3, h: 0.3, fontFace: FONT, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
  s2.addNotes('Wynik: 2/3 budżetu sprzedaży i 2/3 bazy z zamówieniem, ale tylko 21% zamówień. Średnie zamówienie ponad 3× plan, bo klienci robią duże zamówienia na stock. Nowy panel odpowiada na te wnioski: ceny i rabaty liczone od razu w koszyku, kredyt kupiecki zamiast przedpłaty, rabat za krótszy termin płatności, automatyczne maile onboardingowe. Cele FY27: 900 tys. €, 1 125 zamówień.');

  await pres.writeFile({ fileName: 'B2B_PL_kickoff_FY27_2_slajdy.pptx' });
  console.log('ok');
})();
