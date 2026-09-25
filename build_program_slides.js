// 2 slajdy: program partnerski B2B – założenia zweryfikowane z opinią prawną Legal Alliance.
// Styl „1_CEE kick off FY27 ze sprzedażą”. Budowa: NODE_PATH=<node_modules> node build_program_slides.js
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

const C = { teal: '108C96', navy: '174769', ink: '30597A', tint: 'DAEBEA', line: '00676D', ok: '0E8B95', bad: 'C0392B', amber: 'F8AF2D', amberTint: 'FEF3DC', amberTxt: '9A6400', white: 'FFFFFF', muted: '6B7F8E' };
const F = 'Calibri';

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: '#' + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}
const bullets = (items, extra = {}) => items.map((t, j) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: j < items.length - 1, paraSpaceAfter: 5, ...extra } }));
const title = (s, text) => s.addText(text, { x: 0.45, y: 0.3, w: 12.4, h: 0.75, fontFace: F, fontSize: 36, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });
const subtitle = (s, text) => s.addText(text, { x: 0.45, y: 1.02, w: 12.4, h: 0.32, fontFace: F, fontSize: 14, color: C.muted, margin: 0, isTextBox: true });
const header = (s, pres, x, y, w, text) => {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText(text, { x: x + 0.2, y, w: w - 0.3, h: 0.5, fontFace: F, fontSize: 16, bold: true, color: C.white, valign: 'middle', margin: 0, isTextBox: true });
};

async function addSlides(pres) {
  const chk = await icon(fa.FaCheckCircle, C.ok);

  // ================= SLAJD 1: ZAŁOŻENIA =================
  const s1 = pres.addSlide();
  s1.background = { color: C.white };
  title(s1, 'PROGRAM PARTNERSKI B2B – ZAŁOŻENIA');
  subtitle(s1, 'Sklepy ogrodnicze (spółki i JDG) · start 1.10.2026 w wersji ręcznej · zweryfikowane z opinią prawną Legal Alliance');

  // Lewa kolumna: jak działa
  const lx = 0.45, lw = 6.1, ty = 1.55, bh = 4.3;
  header(s1, pres, lx, ty, lw, 'Jak działa');
  s1.addShape(pres.shapes.RECTANGLE, { x: lx, y: ty + 0.5, w: lw, h: bh - 0.5, fill: { color: C.white }, line: { color: C.navy, width: 1.25 } });
  s1.addText([
    { text: '10 pkt', options: { bold: true, color: C.teal, fontSize: 26 } },
    { text: '  za każde 100 zł netto zakupów w panelu', options: { color: C.ink, fontSize: 15 } },
  ], { x: lx + 0.25, y: ty + 0.65, w: lw - 0.5, h: 0.55, fontFace: F, valign: 'middle', margin: 0, isTextBox: true });

  const tiers = [
    { p: '1 000 pkt', a: '≈ 10 tys. zł zakupów', r: 'kody kinowe dla 2 osób (ok. 120 zł)' },
    { p: '10 000 pkt', a: '≈ 100 tys. zł zakupów', r: 'voucher podróżny 5 tys. zł' },
  ];
  const tw = (lw - 0.5 - 0.2) / 2;
  tiers.forEach((t, i) => {
    const x = lx + 0.25 + i * (tw + 0.2), y = ty + 1.35;
    s1.addShape(pres.shapes.RECTANGLE, { x, y, w: tw, h: 1.25, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
    s1.addText(t.p, { x: x + 0.15, y: y + 0.08, w: tw - 0.3, h: 0.45, fontFace: F, fontSize: 20, bold: true, color: C.navy, margin: 0, valign: 'middle', isTextBox: true });
    s1.addText([
      { text: t.a, options: { color: C.muted, fontSize: 11, breakLine: true } },
      { text: t.r, options: { color: C.ink, fontSize: 13, bold: true } },
    ], { x: x + 0.15, y: y + 0.52, w: tw - 0.3, h: 0.68, fontFace: F, margin: 0, valign: 'top', isTextBox: true });
  });
  s1.addText('Każdy, kto osiągnie próg, dostaje nagrodę: bez rankingu i bez losowania.', { x: lx + 0.25, y: ty + 2.72, w: lw - 0.5, h: 0.32, fontFace: F, fontSize: 13, italic: true, color: C.ink, margin: 0, isTextBox: true });
  s1.addText('Punkty bonusowe', { x: lx + 0.25, y: ty + 3.12, w: lw - 0.5, h: 0.3, fontFace: F, fontSize: 14, bold: true, color: C.navy, margin: 0, isTextBox: true });
  s1.addText(bullets([
    'Założenie konta: +500 pkt',
    'Stand w sklepie i post w social media: dobrowolnie, symbolicznie (ok. 100 pkt)',
    'Polecenie sklepu: tylko przy spełnieniu warunków zakupowych',
  ]), { x: lx + 0.25, y: ty + 3.44, w: lw - 0.5, h: 0.8, fontFace: F, fontSize: 12, color: C.ink, valign: 'top', margin: 0, isTextBox: true });

  // Prawa kolumna: potwierdzone przez prawnika
  const rx = 6.75, rw = 6.1;
  header(s1, pres, rx, ty, rw, 'Potwierdzone przez prawnika');
  s1.addShape(pres.shapes.RECTANGLE, { x: rx, y: ty + 0.5, w: rw, h: bh - 0.5, fill: { color: C.white }, line: { color: C.navy, width: 1.25 } });
  const ok = [
    ['Nagroda zawsze dla firmy, nie dla osoby', 'SBM nie jest płatnikiem PIT, nie wystawia PIT-11'],
    ['To nie loteria ani gra hazardowa', 'zezwolenie nie jest potrzebne'],
    ['Podatek od nagrody rozlicza sklep', 'bez dodatkowej nagrody na podatek od SBM'],
    ['Nagrody to koszt marketingowy (CIT)', 'warunek: protokół przekazania każdej nagrody'],
    ['RODO: podstawą regulamin akceptowany przez uczestnika', 'z klauzulą informacyjną'],
    ['CZ/SK w kolejnym sezonie', 'bez zmian podatkowych'],
  ];
  ok.forEach((o, i) => {
    const y = ty + 0.68 + i * 0.6;
    s1.addImage({ data: chk, x: rx + 0.25, y: y + 0.03, w: 0.24, h: 0.24 });
    s1.addText([
      { text: o[0], options: { bold: true, color: C.navy, breakLine: true } },
      { text: o[1], options: { color: C.muted, fontSize: 11.5 } },
    ], { x: rx + 0.6, y, w: rw - 0.8, h: 0.56, fontFace: F, fontSize: 13, valign: 'top', margin: 0, isTextBox: true });
  });

  // Pasek: budżet
  const by = 6.05;
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: by, w: 12.4, h: 0.85, fill: { color: C.navy }, line: { color: C.navy } });
  s1.addText([
    { text: 'BUDŻET NAGRÓD:  ', options: { bold: true, color: C.amber } },
    { text: 'ok. 56 tys. zł na sezon (ok. 13,3 tys. €)  ·  kanał: panel B2B  ·  silnik programu i integracja z panelem w Q4', options: { color: C.white } },
  ], { x: 0.75, y: by, w: 11.9, h: 0.85, fontFace: F, fontSize: 15, valign: 'middle', margin: 0, isTextBox: true });
  s1.addNotes('Mechanika bez zmian: 10 punktów za każde 100 zł netto, próg daje gwarantowaną nagrodę. Prawnik potwierdził, że to nie loteria, że nagroda idzie do firmy, więc nie jesteśmy płatnikiem i nie wystawiamy PIT-11, a sklep sam rozlicza przychód. Nagrody są kosztem marketingowym, pod warunkiem protokołu przekazania. Budżet ok. 56 tys. zł na sezon.');

  // ================= SLAJD 2: CO ZMIENIAMY =================
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  title(s2, 'PROGRAM PARTNERSKI – CO ZMIENIAMY');
  subtitle(s2, 'Zmiany po opinii prawnej Legal Alliance: mniej VAT od nagród i mniejsze ryzyko, że punkty uznane zostaną za zapłatę za usługę');

  const changes = [
    { h: 'Nagroda główna', was: 'Wycieczka w naturze, 5 tys. zł', is: 'Voucher podróżny kwotowy (MPV), 5 tys. zł', why: 'Przy wycieczce powstaje VAT należny od nieodpłatnego przekazania. Voucher MPV jest dla SBM neutralny.' },
    { h: 'Bilety do kina', was: 'Bilety na seans, ok. 120 zł', is: 'Kody lub karty podarunkowe kwotowe (MPV)', why: '120 zł przekracza limit prezentów małej wartości (100 zł netto na osobę rocznie).' },
    { h: 'Stand i post w social media', was: '+500 i +300 pkt', is: 'Aktywność dobrowolna, punkty symboliczne (ok. 100)', why: 'Inaczej punkty to zapłata za usługę marketingową i sklep musi wystawić fakturę.' },
    { h: 'Polecenie sklepu', was: 'Punkty za każdy polecony sklep', is: 'Punkty tylko przy spełnieniu warunków zakupowych', why: 'Inaczej to usługa pośrednictwa, za którą sklep musi wystawić fakturę.' },
  ];
  const cw = 2.95, cg = 0.2, cy = 1.55, ch = 3.25;
  changes.forEach((c, i) => {
    const x = 0.45 + i * (cw + cg);
    header(s2, pres, x, cy, cw, c.h);
    s2.addShape(pres.shapes.RECTANGLE, { x, y: cy + 0.5, w: cw, h: ch - 0.5, fill: { color: C.white }, line: { color: C.navy, width: 1.25 } });
    s2.addText([
      { text: 'Było: ', options: { bold: true, color: C.muted } },
      { text: c.was, options: { color: C.muted, strike: 'sngStrike' } },
    ], { x: x + 0.2, y: cy + 0.65, w: cw - 0.4, h: 0.5, fontFace: F, fontSize: 12.5, valign: 'top', margin: 0, isTextBox: true });
    s2.addText([
      { text: 'Jest: ', options: { bold: true, color: C.teal } },
      { text: c.is, options: { bold: true, color: C.navy } },
    ], { x: x + 0.2, y: cy + 1.2, w: cw - 0.4, h: 0.75, fontFace: F, fontSize: 14, valign: 'top', margin: 0, isTextBox: true });
    s2.addText(c.why, { x: x + 0.2, y: cy + 2.0, w: cw - 0.4, h: 1.15, fontFace: F, fontSize: 11.5, color: C.ink, valign: 'top', margin: 0, isTextBox: true });
  });

  // Do decyzji / do domknięcia
  const dy = 5.0, dh = 1.95, dw = 6.1;
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: dy, w: dw, h: dh, fill: { color: C.amberTint }, line: { color: C.amber, width: 1.25 } });
  s2.addText('DO DECYZJI', { x: 0.7, y: dy + 0.15, w: dw - 0.5, h: 0.32, fontFace: F, fontSize: 15, bold: true, color: C.amberTxt, margin: 0, isTextBox: true });
  s2.addText(bullets([
    'Forma nagród: vouchery MPV (rekomendacja) albo wymiana punktów na rabat na zakupy (najbezpieczniej podatkowo, ale słabiej marketingowo)',
    'Ile punktów za stand i post w social media',
  ]), { x: 0.7, y: dy + 0.55, w: dw - 0.5, h: dh - 0.65, fontFace: F, fontSize: 13, color: C.ink, valign: 'top', margin: 0, isTextBox: true });

  const ox = 6.75;
  s2.addShape(pres.shapes.RECTANGLE, { x: ox, y: dy, w: dw, h: dh, fill: { color: C.navy }, line: { color: C.navy } });
  s2.addText('DO DOMKNIĘCIA Z MECENASEM', { x: ox + 0.25, y: dy + 0.15, w: dw - 0.5, h: 0.32, fontFace: F, fontSize: 15, bold: true, color: C.amber, margin: 0, isTextBox: true });
  s2.addText(bullets([
    'Czy kody kinowe MPV rozwiązują problem limitu 100 zł',
    'Punkty za zgody marketingowe: nie było ich w pytaniach',
    'Odliczenie VAT przy zakupie voucherów MPV (ryzyko sporu)',
    'Przegląd regulaminu przed startem',
  ]), { x: ox + 0.25, y: dy + 0.55, w: dw - 0.5, h: dh - 0.65, fontFace: F, fontSize: 13, color: C.white, valign: 'top', margin: 0, isTextBox: true });
  s2.addNotes('Cztery zmiany po opinii prawnika. Wycieczkę zastępujemy voucherem MPV, bo przy wycieczce w naturze płacimy VAT należny. Bilety do kina zastępujemy kodami kwotowymi, bo 120 zł przekracza limit prezentów małej wartości. Punkty za stand, posty i polecenia mają być dobrowolne i symboliczne, inaczej urząd uzna je za zapłatę za usługę marketingową albo pośrednictwo. Do decyzji: vouchery czy rabat. Do domknięcia z mecenasem: limit 100 zł, punkty za zgody marketingowe, odliczenie VAT przy voucherach i regulamin.');
}

module.exports = { addSlides };

if (require.main === module) {
  (async () => {
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5, jak prezentacja
    pres.title = 'Program partnerski B2B';
    await addSlides(pres);
    await pres.writeFile({ fileName: 'B2B_PL_program_partnerski_2_slajdy.pptx' });
    console.log('ok');
  })();
}
