// ===============================================
// CALCULATORS.GS - Calculation Functions
// ===============================================

function getCourseFromPrice(fullPrice) {
  const price = Number(fullPrice);
  switch (price) {
    case 997: return 'Platinum';
    case 647: return 'Revision';
    case 597: return 'Tuition';
    default: return '';
  }
}

function calculateFMEFee(fullPrice, actualPrice) {
  const validFullPrices = [997, 647, 597];
  const validActualPrices = [997, 647, 597, 397, 347, 297];
  
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