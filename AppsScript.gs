/**
 * Locutores — Web App
 * Extrai dados da planilha incluindo hyperlinks, notas e URLs
 * 
 * COMO USAR:
 * 1. Abra a planilha: Menu → Extensões → Apps Script
 * 2. Cole este código e salve
 * 3. Implante: Implantar → Nova implantação → Web app
 * 4. Copie a URL gerada e cole no index.html
 */

function doGet() {
  const ss = SpreadsheetApp.openById('1Oxf1ooZopUdU2dFtn16lKgENhvggRtzCTbt8UMkMmGk');
  const sheet = ss.getSheets()[0];
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();
  const notes = range.getNotes(); // ← NOVO: captura notas/comentários das células

  const headers = values[0].map(h => String(h).trim());
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowFormulas = formulas[i] || [];
    const rowNotes = notes[i] || [];

    if (!String(row[0] || '').trim()) continue;

    const item = {};

    for (let j = 0; j < headers.length; j++) {
      let value = row[j];
      const formula = String(rowFormulas[j] || '');
      const note = String(rowNotes[j] || '').trim();

      // Priority 1: HYPERLINK formula → extrai URL
      if (formula.startsWith('=HYPERLINK')) {
        const match = formula.match(/"([^"]+)"/);
        if (match) value = match[1];
      }
      // Priority 2: Nota da célula (comentário) → usa como URL
      else if (note && note.startsWith('http')) {
        value = note;
      }
      // Priority 3: Valor já é URL
      else if (String(value).trim().startsWith('http')) {
        value = String(value).trim();
      }

      item[headers[j]] = typeof value === 'string' ? value.trim() : String(value || '');
    }

    result.push(item);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
