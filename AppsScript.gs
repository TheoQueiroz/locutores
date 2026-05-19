/**
 * Locutores — Web App para exportar dados da planilha com hyperlinks reais
 * 
 * Como usar:
 * 1. Abra a planilha: https://docs.google.com/spreadsheets/d/1Oxf1ooZopUdU2dFtn16lKgENhvggRtzCTbt8UMkMmGk
 * 2. Menu: Extensões → Apps Script
 * 3. Cole este código e salve (Ctrl+S)
 * 4. Clique em "Implantar" → "Nova implantação"
 * 5. Tipo: "Web app", Executar como: "Eu", Quem tem acesso: "Qualquer pessoa"
 * 6. Clique em "Implantar" e copie a URL gerada
 * 7. Cole a URL no código do index.html (linha: const APPS_SCRIPT_URL = '...')
 */

function doGet() {
  const ss = SpreadsheetApp.openById('1Oxf1ooZopUdU2dFtn16lKgENhvggRtzCTbt8UMkMmGk');
  const sheet = ss.getSheets()[0];
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();
  
  const headers = values[0].map(h => String(h).trim());
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowFormulas = formulas[i] || [];
    
    if (!String(row[0] || '').trim()) continue;

    const item = {};
    
    for (let j = 0; j < headers.length; j++) {
      let value = row[j];
      
      // Check for HYPERLINK formula
      const formula = String(rowFormulas[j] || '');
      if (formula.startsWith('=HYPERLINK')) {
        const match = formula.match(/"([^"]+)"/);
        value = match ? match[1] : value;
      }
      
      item[headers[j]] = typeof value === 'string' ? value.trim() : String(value || '');
    }
    
    result.push(item);
  }

  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
