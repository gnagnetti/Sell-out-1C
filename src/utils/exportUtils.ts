import { MergedProduct, OrderItem } from '../types';

export function exportSellOutReportCSV(products: MergedProduct[], filename = 'report_sell_out_finale.csv') {
  const headers = [
    'Style Name',
    'Style Clean',
    'Style Number',
    'Color Code',
    'Color Prefix',
    'Silhouette',
    'Linesheet',
    'Fabrication',
    'Image URL',
    'Qty Ordered',
    'Qty Sold',
    'Stock Remaining',
    'Sell Out Pct (%)',
    'Wholesale (EUR)',
    'Sugg. Retail (EUR)',
    'Total Wholesale (EUR)',
    'Realized Retail Revenue (EUR)',
    'PO Numbers',
  ];

  const rows = products.map((p) => [
    `"${(p.styleName || '').replace(/"/g, '""')}"`,
    `"${p.styleClean}"`,
    `"${p.styleNumber}"`,
    `"${(p.colorCode || p.color).replace(/"/g, '""')}"`,
    `"${p.colorPrefix}"`,
    `"${p.silhouette}"`,
    `"${p.linesheet}"`,
    `"${(p.fabrication || '').replace(/"/g, '""')}"`,
    `"${p.image}"`,
    p.unitsOrdered,
    p.unitsSold,
    p.stockRemaining,
    p.sellOutPct.toFixed(1),
    p.wholesaleEur,
    p.suggRetailEur,
    p.totalWholesaleEur,
    p.realizedRevenueEur,
    `"${p.poNumbers.join(', ')}"`,
  ]);

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\n');
  downloadBlob(csvContent, filename, 'text/csv;charset=utf-8;');
}

export function exportConsolidatedOrdersCSV(orders: OrderItem[], filename = 'export_ordini_totale.csv') {
  const headers = [
    'PO Number',
    'Style Name',
    'Style Number',
    'Color',
    'Color Code',
    'Silhouette',
    'Fabrication',
    'Wholesale (EUR)',
    'Sugg. Retail (EUR)',
    'Units Ordered',
    'Total (EUR)',
    'Sizes Breakdown',
    'Status',
  ];

  const rows = orders.map((o) => {
    const sizeStr = Object.entries(o.sizes)
      .map(([sz, q]) => `${sz}:${q}`)
      .join(' | ');

    return [
      `"${o.poNumber}"`,
      `"${(o.styleName || '').replace(/"/g, '""')}"`,
      `"${o.styleNumber}"`,
      `"${o.color}"`,
      `"${(o.colorCode || '').replace(/"/g, '""')}"`,
      `"${o.silhouette}"`,
      `"${(o.fabrication || '').replace(/"/g, '""')}"`,
      o.wholesaleEur,
      o.suggRetailEur,
      o.unitsOrdered,
      o.totalEur,
      `"${sizeStr}"`,
      `"${o.status || 'Approved'}"`,
    ];
  });

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\n');
  downloadBlob(csvContent, filename, 'text/csv;charset=utf-8;');
}

export function downloadSampleSellOutCsv() {
  const content =
`Style Name,Color,Size,Qty Sold
Morolo,Nero,S,1
Morolo,Nero,L,1
Mossa,0319,M,1
Mossa,0319,L,1
Mossa,0319,XL,1
Faggi,3331,40,2
Faggi,3331,42,3
Faggi,3331,44,2
Faggi,3331,46,1
Spettatore,3331,40,1
Spettatore,3331,42,2
Spettatore,3331,44,1
Spettatore,3331,46,1
Domenica,1020,40,2
Domenica,1020,42,3
Domenica,1020,44,1
Domenica,1020,46,1`;

  downloadBlob('\uFEFF' + content, 'template_sell_out_boutique.csv', 'text/csv;charset=utf-8;');
}

function downloadBlob(content: string, filename: string, type: string) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
