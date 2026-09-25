// Slajdy B2B PL w jednym pliku: test panelu (co zagrało / co nie) + wynik i start FY27.
// Budowa: NODE_PATH=<node_modules> node build_all.js
const pptxgen = require('pptxgenjs');
const test = require('./build_test_slides');
const kickoff = require('./build_slides');

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = 'B2B PL – test panelu FY26 i start FY27';
  await test.addSlides(pres);
  await kickoff.addSlides(pres, { skipSummary: true }); // „Podsumowanie FY26” powtarza slajdy testu
  await pres.writeFile({ fileName: 'B2B_PL_slajdy_kickoff_FY27.pptx' });
  console.log('ok');
})();
