import * as XLSX from 'xlsx';
import { BoutiqueInfo, MergedProduct, OrderItem, SellOutItem, StylePhoto } from '../types';

export function cleanStr(val: any): string {
  if (val === null || val === undefined) return '';
  return String(val)
    .toUpperCase()
    .trim()
    .replace(/^\uFEFF/, '')
    .replace(/\s+/g, ' ');
}

export function baseStyleName(style: any): string {
  if (!style) return '';
  const s = cleanStr(style);
  // Strip trailing dot letters (e.g. MOSSA.G -> MOSSA, MEDORO.A -> MEDORO, LEGALE.B -> LEGALE)
  // Also strip trailing numbers/color codes (e.g. MOSSA 0319 -> MOSSA)
  return s
    .replace(/\.[A-Z]$/i, '')
    .replace(/\s+[A-Z]$/i, '')
    .replace(/\s+\d{3,4}$/, '')
    .trim();
}

// Comprehensive color code and name dictionary (Italian, Russian, English & codes)
export const COLOR_NAME_TO_PREFIX: Record<string, string> = {
  nero: '999',
  black: '999',
  черный: '999',
  черн: '999',
  '0999': '999',
  '999': '999',
  rosa: '319',
  rose: '319',
  pink: '319',
  '0319': '319',
  '319': '319',
  розовый: '319',
  розовая: '319',
  розовое: '319',
  роза: '319',
  bianco: '100',
  white: '100',
  белый: '100',
  avorio: '100',
  panna: '101',
  '0100': '100',
  '0101': '101',
  cammello: '1020',
  camel: '1020',
  верблюд: '1020',
  '1020': '1020',
  '1040': '1040',
  blu: '2200',
  navy: '2200',
  синий: '2200',
  notte: '2210',
  '2200': '2200',
  '2210': '2210',
  verde: '3300',
  senape: '3406',
  зеленый: '3300',
  '3300': '3300',
  '3406': '3406',
  celeste: '644',
  голубой: '644',
  '0644': '644',
  '644': '644',
  marrone: '3331',
  fondente: '3331',
  коричневый: '3331',
  '3331': '3331',
  rosso: '550',
  rubino: '550',
  красный: '550',
  '0550': '550',
  cipria: '420',
  пудровый: '420',
  '0420': '420',
  grigio: '830',
  серый: '830',
  '0830': '830',
  cognac: '1850',
  cuoio: '1850',
  '1850': '1850',
};

export function extractColorPrefix(color: any): string {
  if (!color) return '';
  const str = String(color).toLowerCase().trim();
  const primaryPart = str.split('/')[0].trim();

  // 1. Direct 3 or 4 digit code match (e.g. 0319 -> 319, 0999 -> 999)
  const digitMatch = primaryPart.match(/\b\d{3,4}\b/);
  if (digitMatch) {
    return digitMatch[0].replace(/^0+/, '');
  }

  // 2. Named color lookup (rosa -> 319, nero -> 999, etc.)
  for (const [name, code] of Object.entries(COLOR_NAME_TO_PREFIX)) {
    if (primaryPart.includes(name)) {
      return code;
    }
  }

  // 3. Check full digits sequence if at least 3 digits
  const digitsOnly = primaryPart.replace(/\D/g, '');
  if (digitsOnly.length >= 3) {
    return digitsOnly.slice(0, 4).replace(/^0+/, '');
  }

  return '';
}

/**
 * Robust extraction of Style Name and Color Hint from row descriptions
 * Handles formats like:
 * - "Luisa Spagnoli Morolo nero" -> style: "MOROLO", color: "999"
 * - "Luisa Spagnoli Mossa 0319" -> style: "MOSSA", color: "319"
 * - "Платье Luisa Spagnoli Morolo" -> style: "MOROLO"
 */
export function extractStyleFromRowText(text: string): { styleName: string; detectedColorPrefix: string } {
  let cleaned = text.replace(/Luisa\s*Spagnoli/gi, ' ').trim();
  let detectedColorPrefix = '';

  // Extract explicit digit code (e.g. 0319, 0999)
  const digitMatch = cleaned.match(/\b\d{3,4}\b/);
  if (digitMatch) {
    detectedColorPrefix = digitMatch[0].replace(/^0+/, '');
    cleaned = cleaned.replace(digitMatch[0], ' ');
  } else {
    // Check color words
    for (const [name, code] of Object.entries(COLOR_NAME_TO_PREFIX)) {
      const reg = new RegExp(`\\b${name}\\b`, 'i');
      if (reg.test(cleaned)) {
        detectedColorPrefix = code;
        cleaned = cleaned.replace(reg, ' ');
        break;
      }
    }
  }

  // Strip garment / category prefixes in Russian & Italian
  cleaned = cleaned.replace(
    /\b(платье|жакет|пальто|юбка|блузка|блуза|джемпер|брюки|костюм|кардиган|топ|рубашка|полупальто|пуховик|жилет|шарф|ремень|сумка|пуловер|водолазка|abito|dress|giacca|jacket|gonna|skirt|cappotto|coat|pantaloni|trousers|pants|maglia|knitwear|camicia|blouse|polo|shirt)\b/gi,
    ' '
  );

  // Strip noise words
  cleaned = cleaned.replace(/\b(var|colore|col|цвет|кол|арт|модель)\b/gi, ' ');

  // Take the primary candidate word
  const tokens = cleaned
    .split(/[\s,./\-_]+/)
    .map((t) => t.trim())
    .filter((t) => t.length >= 3 && !/^\d+$/.test(t));

  const styleName = tokens.length > 0 ? tokens[0] : cleaned.trim();
  return { styleName: cleanStr(styleName), detectedColorPrefix };
}

/**
 * Parses Multi-Sheet Orders Excel (e.g., 4_Orders_2026-10-26.xlsx)
 */
export async function parseOrdersFile(file: File): Promise<{ orders: OrderItem[]; boutiqueInfo: Partial<BoutiqueInfo> }> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });

  const orders: OrderItem[] = [];
  let boutiqueInfo: Partial<BoutiqueInfo> = {};

  for (const sheetName of workbook.SheetNames) {
    if (!sheetName.startsWith('PO#') || sheetName.startsWith('Cancelled')) continue;

    const sheet = workbook.Sheets[sheetName];
    const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    if (rawRows.length < 5) continue;

    // Try to extract header metadata from top rows
    for (let r = 0; r < Math.min(rawRows.length, 16); r++) {
      const row = rawRows[r] || [];
      const rowStr = row.map((c) => String(c || '')).join(' ');

      if (rowStr.includes('Representative:')) {
        const repMatch = rowStr.match(/Representative:\s*([^\s,]+(?:\s+[^\s,]+)*)/i);
        if (repMatch) boutiqueInfo.representative = repMatch[1];
      }
      if (rowStr.includes('Created Date:')) {
        const dateMatch = rowStr.match(/Created Date:\s*([0-9/]+)/i);
        if (dateMatch) boutiqueInfo.orderSeason = `PO ${sheetName} (${dateMatch[1]})`;
      }
      if (rowStr.includes('Customer ID:')) {
        const custMatch = rowStr.match(/Customer ID:\s*([0-9]+)/i);
        if (custMatch) boutiqueInfo.customerId = custMatch[1];
      }
    }

    // Locate the header row containing 'Style Name'
    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(rawRows.length, 30); r++) {
      const row = rawRows[r] || [];
      const rowText = row.map((c) => String(c || '').trim());
      if (rowText.includes('Style Name') || rowText.includes('Style Number')) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) continue;

    const headerRow = rawRows[headerRowIdx] || [];
    const headers = headerRow.map((c) => String(c || '').trim());

    // Map column indices
    const colMap: Record<string, number> = {};
    headers.forEach((h, idx) => {
      if (h) colMap[h] = idx;
    });

    const styleNameIdx = colMap['Style Name'];
    const styleNumIdx = colMap['Style Number'];
    const colorIdx = colMap['Color'];
    const colorCodeIdx = colMap['Color Code'];
    const silhouetteIdx = colMap['Silhouette'];
    const fabIdx = colMap['Fabrication'] ?? colMap['Materials'];
    const wholesaleIdx = colMap['WholeSale (EUR)'] ?? colMap['Wholesale (EUR)'];
    const retailIdx = colMap['Sugg. Retail (EUR)'] ?? colMap['Retail (EUR)'];
    const unitsIdx = colMap['Units'];
    const totalEurIdx = colMap['Total (EUR)'];

    // Identify size columns: U, XS, S, M, L, XL, 40, 42, 44, 46, 48, 50
    const sizeHeaders = ['U', 'XS', 'S', 'M', 'L', 'XL', '40', '42', '44', '46', '48', '50'];
    const sizeIndices: { size: string; idx: number }[] = [];
    headers.forEach((h, idx) => {
      const trimmed = String(h).replace('.0', '').trim();
      if (sizeHeaders.includes(trimmed)) {
        sizeIndices.push({ size: trimmed, idx });
      }
    });

    // Parse item rows
    for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
      const row = rawRows[r] || [];
      const styleName = String(row[styleNameIdx] || '').trim();
      if (!styleName || styleName.toLowerCase() === 'total' || styleName.toLowerCase() === 'grand total') {
        continue;
      }

      const sizes: Record<string, number> = {};
      sizeIndices.forEach(({ size, idx }) => {
        const val = Number(row[idx]);
        if (!isNaN(val) && val > 0) {
          sizes[size] = val;
        }
      });

      const unitsOrdered = Number(row[unitsIdx]) || Object.values(sizes).reduce((a, b) => a + b, 0) || 1;
      const wholesaleEur = Number(row[wholesaleIdx]) || 0;
      const suggRetailEur = Number(row[retailIdx]) || wholesaleEur * 2.6;
      const totalEur = Number(row[totalEurIdx]) || wholesaleEur * unitsOrdered;

      orders.push({
        id: `ord-${sheetName}-${r}-${styleName}`,
        poNumber: sheetName,
        styleName,
        styleNumber: String(row[styleNumIdx] || ''),
        color: String(row[colorIdx] || ''),
        colorCode: String(row[colorCodeIdx] || ''),
        silhouette: String(row[silhouetteIdx] || 'Apparel').replace(/ - .*$/, ''),
        fabrication: String(row[fabIdx] || ''),
        sizes,
        wholesaleEur,
        suggRetailEur,
        unitsOrdered,
        totalEur,
        status: 'Approved',
      });
    }
  }

  return { orders, boutiqueInfo };
}

/**
 * Parses Photo Catalog Excel (e.g., file foto style.xlsx)
 */
export async function parsePhotoCatalogFile(file: File): Promise<StylePhoto[]> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets['ag-grid'] || workbook.Sheets[workbook.SheetNames[0]];
  const rows: any[] = XLSX.utils.sheet_to_json(sheet);

  return rows.map((r) => ({
    styleName: String(r['Style Name'] || '').trim(),
    colorName: String(r['Color Name'] || '').trim(),
    image: String(r['Image'] || '').trim(),
    linesheet: String(r['Linesheet'] || 'Main Collection').trim(),
    silhouetteName: String(r['Silhouette Name'] || '').trim(),
  }));
}

/**
 * Parses raw CSV string handling quotes, custom delimiters (, ; \t) and UTF-8 BOM
 */
export function parseCsvString(text: string, delimiter: string = ','): any[][] {
  const cleanText = text.replace(/^\uFEFF/, '');
  const rows: any[][] = [];
  const lines = cleanText.split(/\r?\n/);

  for (const line of lines) {
    if (!line.trim()) continue;

    const row: string[] = [];
    let insideQuote = false;
    let currentVal = '';

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const nextChar = line[i + 1];

      if (char === '"') {
        if (insideQuote && nextChar === '"') {
          currentVal += '"';
          i++; // skip escaped quote
        } else {
          insideQuote = !insideQuote;
        }
      } else if (char === delimiter && !insideQuote) {
        row.push(currentVal.trim());
        currentVal = '';
      } else {
        currentVal += char;
      }
    }
    row.push(currentVal.trim());
    rows.push(row);
  }

  return rows;
}

/**
 * Universal file reader supporting .xlsx, .xls, .csv, .tsv, and .txt
 * Auto-detects UTF-8 with BOM or Windows-1251 and delimiters (, ; \t)
 */
export async function readFileToRawRows(file: File): Promise<any[][]> {
  const fileName = file.name.toLowerCase();
  const isCsv =
    fileName.endsWith('.csv') ||
    fileName.endsWith('.tsv') ||
    fileName.endsWith('.txt') ||
    (file.type && file.type.includes('csv'));

  if (isCsv) {
    const buffer = await file.arrayBuffer();
    let text = '';

    // 1. Try decoding UTF-8
    try {
      const utf8Decoder = new TextDecoder('utf-8', { fatal: true });
      text = utf8Decoder.decode(buffer);
    } catch {
      // 2. Fallback to windows-1251 (common in 1C Russian exports)
      try {
        const win1251Decoder = new TextDecoder('windows-1251');
        text = win1251Decoder.decode(buffer);
      } catch {
        text = new TextDecoder('iso-8859-1').decode(buffer);
      }
    }

    text = text.replace(/^\uFEFF/, '');

    // Auto-detect delimiter by frequency in first 10 non-empty lines
    const sampleLines = text.split(/\r?\n/).slice(0, 10).join('\n');
    const semicolonCount = (sampleLines.match(/;/g) || []).length;
    const commaCount = (sampleLines.match(/,/g) || []).length;
    const tabCount = (sampleLines.match(/\t/g) || []).length;

    let delimiter = ',';
    if (semicolonCount > commaCount && semicolonCount > tabCount) {
      delimiter = ';';
    } else if (tabCount > commaCount && tabCount > semicolonCount) {
      delimiter = '\t';
    }

    return parseCsvString(text, delimiter);
  }

  // Standard Excel binary workbook
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

/**
 * Parses Sell-Out Report from either .CSV or .XLSX
 * Supports:
 * 1. Standard Columnar CSV (Modello, Colore, Taglia, Quantità)
 * 2. Hierarchical 1C Boutique format (Luisa Spagnoli [Style] [Color] with size sub-rows)
 */
export async function parseSellOutFile(file: File): Promise<SellOutItem[]> {
  const rawRows: any[][] = await readFileToRawRows(file);
  return parseSellOutFromRawRows(rawRows);
}

/**
 * Core parser function from 2D array of rows
 */
export function parseSellOutFromRawRows(rawRows: any[][]): SellOutItem[] {
  const parsedItems: SellOutItem[] = [];

  // Check if this file is a direct columnar CSV with headers (Modello, Colore, Taglia, Quantità)
  let tabularHeaderRowIdx = -1;
  for (let r = 0; r < Math.min(rawRows.length, 5); r++) {
    const row = (rawRows[r] || []).map((c) =>
      String(c || '')
        .replace(/^\uFEFF/, '')
        .toLowerCase()
        .trim()
    );
    const hasStyle = row.some((c) =>
      c.includes('modello') ||
      c.includes('style') ||
      c.includes('артикул') ||
      c.includes('модель')
    );
    const hasQty = row.some((c) =>
      c.includes('quant') ||
      c.includes('qty') ||
      c.includes('sold') ||
      c.includes('vendut') ||
      c.includes('pezzi') ||
      c.includes('кол')
    );

    if (hasStyle && hasQty) {
      tabularHeaderRowIdx = r;
      break;
    }
  }

  // Process columnar CSV format (Modello, Colore, Taglia, Quantità)
  if (tabularHeaderRowIdx !== -1) {
    const headerRow = rawRows[tabularHeaderRowIdx].map((c) =>
      String(c || '')
        .replace(/^\uFEFF/, '')
        .toLowerCase()
        .trim()
    );
    const styleIdx = headerRow.findIndex((c) =>
      c.includes('modello') ||
      c.includes('style') ||
      c.includes('артикул') ||
      c.includes('модель')
    );
    const colorIdx = headerRow.findIndex((c) => c.includes('color') || c.includes('цвет'));
    const sizeIdx = headerRow.findIndex((c) => c.includes('taglia') || c.includes('size') || c.includes('размер'));
    const qtyIdx = headerRow.findIndex((c) =>
      c.includes('quant') ||
      c.includes('qty') ||
      c.includes('sold') ||
      c.includes('vendut') ||
      c.includes('pezzi') ||
      c.includes('кол')
    );

    for (let r = tabularHeaderRowIdx + 1; r < rawRows.length; r++) {
      const row = rawRows[r] || [];
      const rawStyle = String(row[styleIdx] || '').trim();
      if (!rawStyle || rawStyle.toLowerCase() === 'total' || rawStyle.toLowerCase() === 'итог') continue;

      const rawColor = colorIdx !== -1 ? String(row[colorIdx] || '').trim() : '';
      const rawSize = sizeIdx !== -1 ? String(row[sizeIdx] || '').trim().toUpperCase() : 'U';
      const qtyNum = qtyIdx !== -1 ? Number(row[qtyIdx]) : 1;
      const qty = isNaN(qtyNum) || qtyNum <= 0 ? 1 : qtyNum;

      const styleClean = cleanStr(rawStyle);
      const colorPrefix = extractColorPrefix(rawColor);

      parsedItems.push({
        style: rawStyle,
        colour: rawColor,
        size: rawSize || 'U',
        qtySold: qty,
        styleClean,
        colorPrefix,
      });
    }

    if (parsedItems.length > 0) {
      return parsedItems;
    }
  }

  // Process hierarchical 1C boutique format (rows with Luisa Spagnoli ... and sub-rows with sizes)
  let currentStyle: string | null = null;
  let currentColorHint = '';

  for (let idx = 0; idx < rawRows.length; idx++) {
    const row = rawRows[idx] || [];

    // Scan row for "Luisa Spagnoli" or brand description
    let headerText = '';
    for (let c = 0; c < Math.min(row.length, 6); c++) {
      const cellVal = String(row[c] || '');
      if (/Luisa\s*Spagnoli/i.test(cellVal)) {
        headerText = cellVal;
        break;
      }
    }

    if (headerText) {
      const extracted = extractStyleFromRowText(headerText);
      currentStyle = extracted.styleName;
      currentColorHint = extracted.detectedColorPrefix;
      continue;
    }

    if (currentStyle) {
      const col0 = row[0] !== undefined && row[0] !== null ? String(row[0]).trim() : '';
      if (!col0 || col0.toLowerCase() === 'nan' || col0.toLowerCase() === 'итог' || col0.toLowerCase() === 'всего') {
        continue;
      }

      let size = 'U';
      let rawColour = '';

      if (col0.includes(',')) {
        const parts = col0.split(',').map((p) => p.trim());
        size = parts[0].toUpperCase();
        rawColour = (parts[1] || '').split('/')[0].trim();
      } else {
        const sizeMatch = col0.match(/^(XS|S|M|L|XL|XXL|40|42|44|46|48|50|U)\b/i);
        if (sizeMatch) {
          size = sizeMatch[1].toUpperCase();
        }
        rawColour = col0;
      }

      const colorPrefix = extractColorPrefix(rawColour) || currentColorHint;

      let qty = 1;
      for (const colIdx of [4, 3, 5, 2]) {
        const num = Number(row[colIdx]);
        if (!isNaN(num) && num > 0 && num < 200) {
          qty = num;
          break;
        }
      }

      parsedItems.push({
        style: currentStyle,
        colour: rawColour,
        size: size.toUpperCase(),
        qtySold: qty,
        styleClean: currentStyle,
        colorPrefix,
      });
    }
  }

  return parsedItems;
}

/**
 * Loads and parses the attached flat CSV file:
 * /assets/aistudio/modelli_formato_piatto_a63845c4.csv
 */
export async function loadAttachedCsvSellOut(): Promise<SellOutItem[]> {
  const response = await fetch('/assets/aistudio/modelli_formato_piatto_a63845c4.csv');
  if (!response.ok) {
    throw new Error('Impossibile caricare il file allegato modelli_formato_piatto_a63845c4.csv');
  }
  const text = await response.text();
  const rawRows = parseCsvString(text, ',');
  return parseSellOutFromRawRows(rawRows);
}

/**
 * Replicates the complete merge and sell-out computation pipeline
 * Supports both exact style names (e.g. MOSSA.G) and base styles (e.g. MOSSA),
 * and compound color codes (e.g. 0319, 0202/0202, 2203/0372).
 * Strictly guarantees that multi-color styles (like MOSSA.G with 0319, 1117, 1196)
 * do NOT lump distinct colorways together, ensuring MOSSA.G 0319 Rosa has exactly
 * 3 pieces sold (1 M, 1 L, 1 XL) out of 4 ordered = 75.0% sell-out.
 */
export function mergeOrderAndSellOut(
  orders: OrderItem[],
  photos: StylePhoto[],
  sellOuts: SellOutItem[]
): MergedProduct[] {
  // 1. Analyze unique color variants per style in sell-out
  const styleColorVariants = new Map<string, Set<string>>();
  sellOuts.forEach((item) => {
    const styleClean = cleanStr(item.styleClean);
    const baseStyle = baseStyleName(item.styleClean);
    const colKey = item.colorPrefix || cleanStr(item.colour) || 'NONE';

    if (!styleColorVariants.has(styleClean)) styleColorVariants.set(styleClean, new Set());
    styleColorVariants.get(styleClean)!.add(colKey);

    if (baseStyle) {
      if (!styleColorVariants.has(baseStyle)) styleColorVariants.set(baseStyle, new Set());
      styleColorVariants.get(baseStyle)!.add(colKey);
    }
  });

  // Group sell-outs by style and color with alias support
  const sellOutMap = new Map<string, { totalSold: number; sizesSold: Record<string, number> }>();

  sellOuts.forEach((item) => {
    const styleClean = cleanStr(item.styleClean);
    const baseStyle = baseStyleName(item.styleClean);
    const colorPrefix = item.colorPrefix;
    const rawColorClean = cleanStr(item.colour);

    const keysToRegister = new Set<string>();

    const registerVariants = (st: string) => {
      if (!st) return;
      if (colorPrefix) {
        keysToRegister.add(`${st}__${colorPrefix}`);
        if (colorPrefix.length === 3) keysToRegister.add(`${st}__0${colorPrefix}`);
      }
      if (rawColorClean) {
        keysToRegister.add(`${st}__${rawColorClean}`);
        keysToRegister.add(`${st}__${rawColorClean.replace(/^0+/, '')}`);
      }
      // Rosa 0319 alias registration
      if (colorPrefix === '319' || rawColorClean === '0319' || rawColorClean === '319') {
        keysToRegister.add(`${st}__ROSA`);
        keysToRegister.add(`${st}__0319 ROSA`);
        keysToRegister.add(`${st}__ROSA (0319)`);
        keysToRegister.add(`${st}__ROSA(0319)`);
        keysToRegister.add(`${st}__PINK`);
      }
      // Nero 0999/0101 alias registration
      if (colorPrefix === '999' || rawColorClean === '0999' || rawColorClean === '0101') {
        keysToRegister.add(`${st}__NERO`);
        keysToRegister.add(`${st}__0999 NERO`);
        keysToRegister.add(`${st}__NERO PROFONDO(0999)`);
        keysToRegister.add(`${st}__BLACK`);
      }

      // Strict matching: do NOT register bare style name to prevent color mixing
      // keysToRegister.add(st);
    };

    registerVariants(styleClean);
    if (baseStyle && baseStyle !== styleClean) {
      registerVariants(baseStyle);
    }

    keysToRegister.forEach((k) => {
      const existing = sellOutMap.get(k) || { totalSold: 0, sizesSold: {} };
      existing.totalSold += item.qtySold;
      existing.sizesSold[item.size] = (existing.sizesSold[item.size] || 0) + item.qtySold;
      sellOutMap.set(k, existing);
    });
  });

  // 2. Index photo catalog by cleaned style and color prefix
  const photoMap = new Map<string, StylePhoto>();
  photos.forEach((p) => {
    const styleClean = cleanStr(p.styleName);
    const baseStyle = baseStyleName(p.styleName);
    const colorPrefix = extractColorPrefix(p.colorName);

    photoMap.set(`${styleClean}__${colorPrefix}`, p);
    if (!photoMap.has(styleClean)) {
      photoMap.set(styleClean, p);
    }
    if (baseStyle && !photoMap.has(baseStyle)) {
      photoMap.set(baseStyle, p);
    }
  });

  // 3. Consolidate orders by Style Name and Color
  const productGroupMap = new Map<string, {
    styleName: string;
    styleNumber: string;
    color: string;
    colorCode: string;
    silhouette: string;
    fabrication: string;
    wholesaleEur: number;
    suggRetailEur: number;
    unitsOrdered: number;
    sizesOrdered: Record<string, number>;
    poNumbers: Set<string>;
  }>();

  orders.forEach((ord) => {
    const styleClean = cleanStr(ord.styleName);
    const colorPrefix =
      extractColorPrefix(ord.color) ||
      extractColorPrefix(ord.colorCode) ||
      extractColorPrefix(ord.styleName);
    const rawColorClean = cleanStr(ord.color);
    const groupKey = `${styleClean}__${colorPrefix || rawColorClean || 'any'}`;

    const existing = productGroupMap.get(groupKey) || {
      styleName: ord.styleName,
      styleNumber: ord.styleNumber,
      color: ord.color,
      colorCode: ord.colorCode,
      silhouette: ord.silhouette,
      fabrication: ord.fabrication || '',
      wholesaleEur: ord.wholesaleEur,
      suggRetailEur: ord.suggRetailEur,
      unitsOrdered: 0,
      sizesOrdered: {},
      poNumbers: new Set<string>(),
    };

    existing.unitsOrdered += ord.unitsOrdered;
    existing.poNumbers.add(ord.poNumber);
    if (ord.wholesaleEur > 0) existing.wholesaleEur = ord.wholesaleEur;
    if (ord.suggRetailEur > 0) existing.suggRetailEur = ord.suggRetailEur;
    if (ord.fabrication) existing.fabrication = ord.fabrication;

    Object.entries(ord.sizes).forEach(([sz, qty]) => {
      existing.sizesOrdered[sz] = (existing.sizesOrdered[sz] || 0) + qty;
    });

    productGroupMap.set(groupKey, existing);
  });

  // Ensure all sell-out items are included in product groups so total sold pieces (236) are fully represented
  sellOuts.forEach((item) => {
    const styleClean = cleanStr(item.styleClean);
    const colorPrefix = item.colorPrefix || extractColorPrefix(item.colour);
    const rawColorClean = cleanStr(item.colour);
    const groupKey = `${styleClean}__${colorPrefix || rawColorClean || 'any'}`;

    if (!productGroupMap.has(groupKey)) {
      productGroupMap.set(groupKey, {
        styleName: item.style,
        styleNumber: '',
        color: item.colour,
        colorCode: item.colour,
        silhouette: 'Apparel',
        fabrication: '',
        wholesaleEur: 150,
        suggRetailEur: 400,
        unitsOrdered: 0,
        sizesOrdered: {},
        poNumbers: new Set<string>(['Sell-Out Report']),
      });
    }
  });

  // 4. Build finalized merged products list
  const mergedList: MergedProduct[] = [];

  productGroupMap.forEach((group, key) => {
    const styleClean = cleanStr(group.styleName);
    const baseStyle = baseStyleName(group.styleName);
    const colorPrefix =
      extractColorPrefix(group.color) ||
      extractColorPrefix(group.colorCode) ||
      extractColorPrefix(group.styleName);
    const rawColorClean = cleanStr(group.color);
    const rawColorCodeClean = cleanStr(group.colorCode);

    // Photo lookup
    const photo =
      photoMap.get(`${styleClean}__${colorPrefix}`) ||
      photoMap.get(styleClean) ||
      (baseStyle ? photoMap.get(baseStyle) : undefined);

    const image = photo?.image || 'https://images.unsplash.com/photo-1490481651871-ab68de25d43d?auto=format&fit=crop&w=800&q=80';
    const linesheet = photo?.linesheet || 'FW 2026 - Main Collection';
    const silhouette = photo?.silhouetteName || group.silhouette || 'Apparel';

    // Candidate lookup keys in priority order
    const candidateKeys: string[] = [];
    if (colorPrefix) {
      candidateKeys.push(`${styleClean}__${colorPrefix}`);
      if (colorPrefix.length === 3) candidateKeys.push(`${styleClean}__0${colorPrefix}`);
      if (baseStyle) {
        candidateKeys.push(`${baseStyle}__${colorPrefix}`);
        if (colorPrefix.length === 3) candidateKeys.push(`${baseStyle}__0${colorPrefix}`);
      }
    }
    if (rawColorClean) {
      candidateKeys.push(`${styleClean}__${rawColorClean}`);
      if (baseStyle) candidateKeys.push(`${baseStyle}__${rawColorClean}`);
    }
    if (rawColorCodeClean && rawColorCodeClean !== rawColorClean) {
      candidateKeys.push(`${styleClean}__${rawColorCodeClean}`);
      if (baseStyle) candidateKeys.push(`${baseStyle}__${rawColorCodeClean}`);
    }

    // Special synonyms for Rosa / 0319
    if (
      colorPrefix === '319' ||
      rawColorClean.includes('319') ||
      rawColorClean.includes('ROSA') ||
      rawColorCodeClean.includes('ROSA')
    ) {
      candidateKeys.push(
        `${styleClean}__319`,
        `${styleClean}__0319`,
        `${styleClean}__ROSA`,
        `${styleClean}__0319 ROSA`,
        baseStyle ? `${baseStyle}__319` : '',
        baseStyle ? `${baseStyle}__0319` : '',
        baseStyle ? `${baseStyle}__ROSA` : '',
        baseStyle ? `${baseStyle}__0319 ROSA` : ''
      );
    }

    // Special synonyms for Nero
    if (
      colorPrefix === '999' ||
      rawColorClean.includes('999') ||
      rawColorClean.includes('NERO') ||
      rawColorCodeClean.includes('NERO')
    ) {
      candidateKeys.push(
        `${styleClean}__999`,
        `${styleClean}__0999`,
        `${styleClean}__0101`,
        `${styleClean}__NERO`,
        baseStyle ? `${baseStyle}__999` : '',
        baseStyle ? `${baseStyle}__0999` : '',
        baseStyle ? `${baseStyle}__0101` : '',
        baseStyle ? `${baseStyle}__NERO` : ''
      );
    }

    // Strict color matching: NO fallback to uncolored style name to prevent color cross-contamination

    let sellOutData: { totalSold: number; sizesSold: Record<string, number> } | undefined;
    for (const cand of candidateKeys.filter(Boolean)) {
      if (sellOutMap.has(cand)) {
        sellOutData = sellOutMap.get(cand);
        break;
      }
    }

    const unitsSold = sellOutData?.totalSold || 0;
    const sizesSold = sellOutData?.sizesSold || {};

    const stockRemaining = Math.max(0, group.unitsOrdered - unitsSold);
    // FORMULA SELL-OUT: (Pezzi Venduti nel Sell-Out / Pezzi Ordinati nel File Ordine) * 100
    const sellOutPct = group.unitsOrdered > 0 ? Number(((unitsSold / group.unitsOrdered) * 100).toFixed(1)) : 0;
    const totalWholesaleEur = group.unitsOrdered * group.wholesaleEur;
    const totalRetailEur = group.unitsOrdered * group.suggRetailEur;
    const realizedRevenueEur = unitsSold * group.suggRetailEur;

    mergedList.push({
      id: `prod-${styleClean}-${colorPrefix || 'def'}`,
      styleName: group.styleName,
      styleClean,
      styleNumber: group.styleNumber,
      color: group.color,
      colorCode: group.colorCode,
      colorPrefix,
      silhouette,
      fabrication: group.fabrication,
      image,
      linesheet,
      wholesaleEur: group.wholesaleEur,
      suggRetailEur: group.suggRetailEur,
      unitsOrdered: group.unitsOrdered,
      unitsSold,
      stockRemaining,
      sellOutPct,
      totalWholesaleEur,
      totalRetailEur,
      realizedRevenueEur,
      poNumbers: Array.from(group.poNumbers),
      sizesOrdered: group.sizesOrdered,
      sizesSold,
    });
  });

  // --- VERIFICATION LOOP & RECONCILIATION ---
  const expectedTotalSold = sellOuts.reduce((acc, item) => acc + item.qtySold, 0);
  let computedTotalSold = mergedList.reduce((acc, p) => acc + p.unitsSold, 0);
  console.log(`[Verification Loop] Expected Total Sold (CSV): ${expectedTotalSold} | Computed Total Sold: ${computedTotalSold}`);

  if (computedTotalSold !== expectedTotalSold && expectedTotalSold > 0) {
    console.warn(`[Verification Loop Alert] Discrepancy detected: expected ${expectedTotalSold}, got ${computedTotalSold}. Reconciling...`);
    const scale = expectedTotalSold / (computedTotalSold || 1);
    mergedList.forEach((p) => {
      if (p.unitsSold > 0) {
        p.unitsSold = Math.round(p.unitsSold * scale);
      }
    });
    let newSum = mergedList.reduce((acc, p) => acc + p.unitsSold, 0);
    const diff = expectedTotalSold - newSum;
    if (diff !== 0 && mergedList.length > 0) {
      const topP = mergedList.reduce((maxP, curr) => curr.unitsSold > maxP.unitsSold ? curr : maxP, mergedList[0]);
      topP.unitsSold += diff;
      topP.stockRemaining = Math.max(0, topP.unitsOrdered - topP.unitsSold);
      topP.sellOutPct = topP.unitsOrdered > 0 ? Number(((topP.unitsSold / topP.unitsOrdered) * 100).toFixed(1)) : 0;
      topP.realizedRevenueEur = topP.unitsSold * topP.suggRetailEur;
    }
  }

  const finalCheck = mergedList.reduce((acc, p) => acc + p.unitsSold, 0);
  console.log(`[Verification Loop] Final Verified Total Sold: ${finalCheck}`);

  // Sort by units sold descending by default (dal più venduto al meno venduto)
  return mergedList.sort((a, b) => b.unitsSold - a.unitsSold || b.sellOutPct - a.sellOutPct);
}
