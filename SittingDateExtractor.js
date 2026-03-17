// ===============================================
// SITTING DATE EXTRACTOR
// Extracts just the date from full Item Name
// ===============================================

function extractSittingDate(fullItemName) {
  if (!fullItemName || fullItemName === '') {
    Logger.log('⚠️ No sitting data provided');
    return '';
  }
  
  const fullText = fullItemName.toString().trim();
  
  // Split on " - " (space-dash-space)
  const parts = fullText.split(' - ');
  
  if (parts.length >= 2) {
    // Take the last part (the date)
    const sittingDate = parts[parts.length - 1].trim();
    Logger.log(`✅ Extracted sitting date: "${fullText}" → "${sittingDate}"`);
    return sittingDate;
  }
  
  // Fallback: if no dash found, return the original
  // (This handles edge cases where format is different)
  Logger.log(`⚠️ No dash found in sitting, using full text: "${fullText}"`);
  return fullText;
}