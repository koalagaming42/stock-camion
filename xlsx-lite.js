/* Générateur Excel (.xlsx) minimal, sans dépendance, pour l'appli Stock Camion.
   Produit un classeur mis en forme : en-têtes, lignes « À commander » en rouge, filtres,
   volets figés, onglet Historique. Fonctionne hors ligne. */
(function (global) {
  'use strict';

  // ---------- ZIP (sans compression) ----------
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) { let c = 0xFFFFFFFF; for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  const enc = new TextEncoder();
  function zip(files) { // files: [{name, data(string)}]
    const parts = [], central = []; let offset = 0;
    const d = new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
          date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const f of files) {
      const name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, time, true); h.setUint16(12, date, true); h.setUint32(14, crc, true);
      h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
      c.setUint16(12, time, true); c.setUint16(14, date, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint16(30, 0, true); c.setUint16(32, 0, true); c.setUint16(34, 0, true); c.setUint16(36, 0, true);
      c.setUint32(38, 0, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((s, p) => s + p.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(e.buffer)];
    const out = new Uint8Array(all.reduce((s, p) => s + p.length, 0)); let o = 0;
    for (const p of all) { out.set(p, o); o += p.length; }
    return out;
  }

  // ---------- Feuilles ----------
  const xe = s => String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colName = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };
  function cell(ref, v, style) {
    if (v === null || v === undefined || v === '') return style ? `<c r="${ref}" s="${style}"/>` : '';
    if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}" s="${style || 0}"><v>${v}</v></c>`;
    return `<c r="${ref}" s="${style || 0}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
  }
  // rows: [[{v, s}]] ; opts: {widths, freeze, filter}
  function sheetXml(rows, opts) {
    const cols = (opts.widths || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
    const body = rows.map((r, ri) => `<row r="${ri + 1}">` + r.map((c, ci) => c ? cell(colName(ci) + (ri + 1), c.v, c.s) : '').join('') + '</row>').join('');
    const pane = opts.freeze ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : '';
    const af = opts.filter ? `<autoFilter ref="${opts.filter}"/>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${pane}<sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${body}</sheetData>${af}</worksheet>`;
  }

  // Styles : 0 normal · 1 en-tête · 2 texte · 3 nombre centré · 4 texte rouge · 5 nombre rouge · 6 titre résumé · 7 texte gras · 8 nombre gras centré
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6">
<font><sz val="11"/><name val="Arial"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
<font><sz val="11"/><color rgb="FF0000FF"/><name val="Arial"/></font>
<font><b/><sz val="11"/><color rgb="FFC62828"/><name val="Arial"/></font>
<font><b/><sz val="12"/><color rgb="FF1F3A5F"/><name val="Arial"/></font>
<font><b/><sz val="11"/><name val="Arial"/></font>
</fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3A5F"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFDE8E8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFD0D7E2"/></left><right style="thin"><color rgb="FFD0D7E2"/></right><top style="thin"><color rgb="FFD0D7E2"/></top><bottom style="thin"><color rgb="FFD0D7E2"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="9">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  function workbook(sheets) { // sheets: [{name, xml}]
    const files = [];
    files.push({ name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>` });
    files.push({ name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` });
    files.push({ name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xe(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${sheets.some(s => s.filter) ? `<definedNames>${sheets.map((s, i) => s.filter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xe(s.name)}'!${s.filter.split(':').map(x => x.replace(/([A-Z]+)(\d+)/, '$$$1$$$2')).join(':')}</definedName>` : '').join('')}</definedNames>` : ''}</workbook>` });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` });
    files.push({ name: 'xl/styles.xml', data: STYLES });
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: s.xml }));
    return zip(files);
  }

  /* Construit le classeur Stock Camion.
     items: [{name, ref, cat, qty, min, target, ordered}] ; history: [{t, name, delta, after, note}] ; extra: {pending(id)} */
  function stockWorkbook(items, history, opts) {
    opts = opts || {};
    const pending = opts.pending || (() => 0);
    const num = v => (v === '' || v == null || !isFinite(+v)) ? '' : +v;
    const low = it => it.qty <= it.min;
    const toOrder = it => low(it) ? Math.max((it.target || it.min) - it.qty, 1) : 0;
    const status = it => pending(it) ? 'Commandé' : (low(it) ? 'À commander' : 'OK');
    const hdr = ['Désignation', 'Référence', 'Catégorie', 'Stock', 'Seuil', 'Cible', 'Statut', 'À commander'];
    const rows = [hdr.map(v => ({ v, s: 1 }))];
    items.forEach(it => {
      const red = low(it);
      const t = red ? 4 : 2, n = red ? 5 : 3;
      rows.push([{ v: it.name, s: t }, { v: it.ref || '', s: t }, { v: it.cat || '', s: t },
        { v: num(it.qty), s: n }, { v: num(it.min), s: n }, { v: num(it.target), s: n },
        { v: status(it), s: n }, { v: toOrder(it), s: n }]);
    });
    const nLow = items.filter(low).length, nPieces = items.reduce((s, it) => s + toOrder(it), 0);
    const side = [[{ v: 'Résumé', s: 6 }], [{ v: 'Articles à commander', s: 7 }, { v: nLow, s: 8 }],
      [{ v: 'Pièces à commander', s: 7 }, { v: nPieces, s: 8 }], [{ v: 'Mis à jour le', s: 7 }, { v: opts.date || '', s: 0 }]];
    side.forEach((r, i) => { while (rows.length <= i) rows.push([]); const row = rows[i]; while (row.length < 9) row.push(null); row[9] = r[0]; row[10] = r[1] || null; });
    const last = items.length + 1;
    const s1 = { name: 'Stock camion', filter: `A1:H${last}`, xml: sheetXml(rows, { widths: [46, 14, 15, 9, 9, 9, 14, 13, 3, 24, 18], freeze: true, filter: `A1:H${last}` }) };

    const hrows = [['Date', 'Article', 'Mouvement', 'Reste', 'Note'].map(v => ({ v, s: 1 }))];
    (history || []).slice(0, 2000).forEach(h => {
      const out = h.delta < 0;
      hrows.push([{ v: new Date(h.t).toLocaleString('fr-FR'), s: 2 }, { v: h.name, s: 2 },
        { v: h.delta, s: out ? 5 : 3 }, { v: h.after, s: 3 }, { v: h.note || '', s: 2 }]);
    });
    const hl = hrows.length;
    const s2 = { name: 'Historique', filter: `A1:E${hl}`, xml: sheetXml(hrows, { widths: [20, 46, 12, 9, 36], freeze: true, filter: `A1:E${hl}` }) };
    return workbook([s1, s2]);
  }

  global.XlsxLite = { stockWorkbook, zip, crc32 };
})(typeof window !== 'undefined' ? window : globalThis);
