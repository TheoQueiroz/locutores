/**
 * Locutores — Web App para exportar dados da planilha com links reais
 * Usa getRichTextValues() para capturar hyperlinks corretamente
 */
function doGet() {
  const ss = SpreadsheetApp.openById('1Oxf1ooZopUdU2dFtn16lKgENhvggRtzCTbt8UMkMmGk');
  const sheet = ss.getSheets()[0];
  const range = sheet.getDataRange();
  const values = range.getValues();
  const richText = range.getRichTextValues();
  
  const headers = values[0].map(h => String(h).trim());
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowRT = richText[i] || [];
    
    if (!String(row[0] || '').trim()) continue;

    const item = {};
    
    for (let j = 0; j < headers.length; j++) {
      let value = row[j];
      const rt = rowRT[j];
      
      // getLinkUrl() returns the hyperlink URL if the cell has one
      if (rt) {
        const linkUrl = rt.getLinkUrl();
        if (linkUrl) {
          value = linkUrl;
        }
      }
      
      item[headers[j]] = typeof value === 'string' ? value.trim() : String(value || '');
    }
    
    result.push(item);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
