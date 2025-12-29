// ===============================================
// CALCULATORS.GS - Calculation Functions
// UPDATED: Platinum course now £1047 (was £997)
// UPDATED: Added Tuition/Revision Plus (822) support
// ===============================================

function getCourseFromPrice(fullPrice) {
  const price = Number(fullPrice);
  switch (price) {
    case 1047: return 'Platinum';  // UPDATED from 997
    case 822: return 'Tuition/Revision Plus';
    case 647: return 'Revision';
    case 597: return 'Tuition';
    default: return '';
  }
}

function calculateFMEFee(fullPrice, actualPrice) {
  // UPDATED: Added 1047 and 350 to valid prices
  const validFullPrices = [1047, 822, 647, 597];
  const validActualPrices = [1047, 822, 647, 597, 522, 397, 350, 347, 300, 297];

  if (validFullPrices.includes(Number(fullPrice)) &&
      validActualPrices.includes(Number(actualPrice))) {
    return Number(fullPrice) * 0.1;
  }

  return '';
}

function calculateStripeFee(actualPrice) {
  return Number(actualPrice) * 0.01;
}

function calculateExpectedIncome(actualPrice, fmeFee, stripeFee) {
  const actual = Number(actualPrice);
  const fme = Number(fmeFee) || 0;
  const stripe = Number(stripeFee) || 0;

  return actual - fme - stripe;
}