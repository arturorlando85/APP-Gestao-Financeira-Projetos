/* Safe Harbor — motor de leitura/escrita da planilha V7 (sem dependências além de JSZip).
   Funciona no navegador (window.SHCore) e no Node (module.exports). */
(function (root) {
  'use strict';

  // ---------- utilidades ----------
  const decodeXml = s => s.replace(/&(lt|gt|quot|apos|amp|#x([0-9a-fA-F]+)|#(\d+));/g, (m, n, hx, dc) =>
    n === 'lt' ? '<' : n === 'gt' ? '>' : n === 'quot' ? '"' : n === 'apos' ? "'" : n === 'amp' ? '&' :
    hx ? String.fromCodePoint(parseInt(hx, 16)) : String.fromCodePoint(parseInt(dc, 10)));
  const encXml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const colToNum = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const numToCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const splitRef = r => { const m = /^([A-Z]+)(\d+)$/.exec(r); return { col: m[1], row: +m[2], c: colToNum(m[1]) }; };
  const norm = s => String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase();
  const EPOCH = Date.UTC(1899, 11, 30);
  const serialToISO = n => { const d = new Date(EPOCH + Math.round(n) * 86400000); return d.toISOString().slice(0, 10); };
  const isoToSerial = s => { const [y, m, d] = s.split('-').map(Number); return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000); };
  const isISO = v => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v);
  const attr = (a, name) => { const m = new RegExp('(?:^|\\s)' + name + '="([^"]*)"').exec(a); return m ? m[1] : null; };

  // ---------- fórmulas: deslocamento de referências relativas ----------
  function shiftFormula(f, dr, dc) {
    if (!dr && !dc) return f;
    let out = '', seg = '', i = 0;
    const flush = () => {
      out += seg.replace(/(^|[^A-Za-z0-9_.$\]])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(\[])/g, (m, pre, d1, col, d2, row) => {
        const c = d1 ? col : numToCol(colToNum(col) + dc);
        const r = d2 ? row : String(+row + dr);
        return pre + d1 + c + d2 + r;
      });
      seg = '';
    };
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"' || ch === "'") {
        flush(); let j = i + 1;
        while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; }
        out += f.slice(i, j + 1); i = j + 1;
      } else if (ch === '[') {
        flush(); let depth = 0, j = i;
        for (; j < f.length; j++) { if (f[j] === '[') depth++; else if (f[j] === ']') { depth--; if (!depth) break; } }
        out += f.slice(i, j + 1); i = j + 1;
      } else { seg += ch; i++; }
    }
    flush();
    return out;
  }

  // ---------- planilha (XML) ----------
  function parseSST(xml) {
    if (!xml) return [];
    const out = [];
    const re = /<si>([\s\S]*?)<\/si>/g; let m;
    while ((m = re.exec(xml))) {
      const body = m[1].replace(/<rPh[\s\S]*?<\/rPh>/g, '');
      let t = ''; const tr = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t\s*\/>/g; let k;
      while ((k = tr.exec(body))) t += k[1] ? decodeXml(k[1]) : '';
      out.push(t);
    }
    return out;
  }

  // Lê as linhas de uma aba preservando o XML bruto de cada célula
  function parseSheet(xml, sst) {
    const sdStart = xml.indexOf('<sheetData');
    const sdOpenEnd = xml.indexOf('>', sdStart) + 1;
    const selfClosed = xml[sdOpenEnd - 2] === '/';
    const sdEnd = selfClosed ? sdOpenEnd : xml.indexOf('</sheetData>');
    const head = xml.slice(0, sdStart), tail = xml.slice(selfClosed ? sdOpenEnd : sdEnd + 12);
    const body = selfClosed ? '' : xml.slice(sdOpenEnd, sdEnd);
    const rows = new Map();
    const rowRe = /<row\s([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g; let rm;
    const shared = {};
    while ((rm = rowRe.exec(body))) {
      const rAttrs = rm[1]; const r = +attr(rAttrs, 'r');
      const cells = new Map();
      const inner = rm[3] || '';
      const cRe = /<c\s([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g; let cm;
      while ((cm = cRe.exec(inner))) {
        const a = cm[1], ci = cm[3] || '';
        const ref = attr(a, 'r'); const { c } = splitRef(ref);
        const cell = { ref, s: attr(a, 's'), t: attr(a, 't'), cm: attr(a, 'cm'), f: null, fAttrs: null, v: null };
        const fm = /<f(\s[^>]*?)?(?:\/>|>([\s\S]*?)<\/f>)/.exec(ci);
        if (fm) {
          cell.fAttrs = fm[1] || '';
          cell.f = fm[2] != null ? decodeXml(fm[2]) : null;
          const ft = attr(cell.fAttrs, 't');
          if (ft === 'shared') {
            const si = attr(cell.fAttrs, 'si');
            if (cell.f != null && cell.f !== '') shared[si] = { f: cell.f, ref };
            else cell.sharedSi = si;
          }
        }
        const vm = /<v>([\s\S]*?)<\/v>/.exec(ci);
        const im = /<is>([\s\S]*?)<\/is>/.exec(ci);
        if (cell.t === 's' && vm) cell.v = sst[+vm[1]];
        else if (cell.t === 'inlineStr' && im) { let t = ''; im[1].replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (x, y) => { t += decodeXml(y); }); cell.v = t; }
        else if (cell.t === 'str' && vm) cell.v = decodeXml(vm[1]);
        else if (cell.t === 'b' && vm) cell.v = vm[1] === '1';
        else if (cell.t === 'e' && vm) cell.v = null, cell.err = vm[1];
        else if (vm) cell.v = +vm[1];
        cells.set(c, cell);
      }
      rows.set(r, { attrs: rAttrs, cells });
    }
    // expande fórmulas compartilhadas
    for (const row of rows.values()) for (const cell of row.cells.values()) {
      if (cell.sharedSi != null) {
        const m = shared[cell.sharedSi];
        if (m) { const a = splitRef(m.ref), b = splitRef(cell.ref); cell.f = shiftFormula(m.f, b.row - a.row, b.c - a.c); }
      }
      if (cell.fAttrs && attr(cell.fAttrs, 't') === 'shared') cell.fAttrs = '';
    }
    return { head, tail, rows };
  }

  function cellXml(ref, s, val, opts = {}) {
    const sa = s != null ? ` s="${s}"` : '';
    const cma = opts.cm ? ` cm="${opts.cm}"` : '';
    if (opts.f != null) {
      let fa = opts.fAttrs || '';
      if (attr(fa, 't') === 'array') fa = fa.replace(/ref="[^"]*"/, `ref="${ref}"`);
      let vx = '', ta = '';
      if (typeof opts.v === 'number' && isFinite(opts.v)) vx = `<v>${opts.v}</v>`;
      else if (typeof opts.v === 'string') { ta = ' t="str"'; vx = `<v>${encXml(opts.v)}</v>`; }
      else if (typeof opts.v === 'boolean') { ta = ' t="b"'; vx = `<v>${opts.v ? 1 : 0}</v>`; }
      else if (opts.err) { ta = ' t="e"'; vx = `<v>${encXml(opts.err)}</v>`; }
      return `<c r="${ref}"${sa}${ta}${cma}><f${fa}>${encXml(opts.f)}</f>${vx}</c>`;
    }
    if (val == null || val === '') return `<c r="${ref}"${sa}/>`;
    if (typeof val === 'number' && isFinite(val)) return `<c r="${ref}"${sa}><v>${val}</v></c>`;
    if (typeof val === 'boolean') return `<c r="${ref}"${sa} t="b"><v>${val ? 1 : 0}</v></c>`;
    if (isISO(val) && opts.date) return `<c r="${ref}"${sa}><v>${isoToSerial(val)}</v></c>`;
    return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${encXml(val)}</t></is></c>`;
  }
  // célula existente sem alteração (fórmula expandida, sem valor em cache)
  function keepCellXml(cell, ref, dr) {
    if (cell.f != null) {
      if (/#REF!/.test(cell.f)) return `<c r="${ref}"${cell.s != null ? ` s="${cell.s}"` : ''}/>`;
      return cellXml(ref, cell.s, null, { f: shiftFormula(cell.f, dr, 0), fAttrs: cell.fAttrs, cm: cell.cm, v: cell.v, err: cell.err });
    }
    return cellXml(ref, cell.s, cell.v, { date: false });
  }
  function serializeRows(rows) {
    const keys = [...rows.keys()].sort((a, b) => a - b);
    let out = '';
    for (const r of keys) {
      const row = rows.get(r);
      const cells = [...row.cells.entries()].sort((a, b) => a[0] - b[0]);
      const a = row.attrs.replace(/(^|\s)r="\d+"/, `$1r="${r}"`);
      if (!cells.length) { out += `<row ${a}/>`; continue; }
      out += `<row ${a}>` + cells.map(([, x]) => x.xml != null ? x.xml : keepCellXml(x, x.ref, 0)).join('') + '</row>';
    }
    return out;
  }

  const cellVal = (rows, r, c) => { const row = rows.get(r); const x = row && row.cells.get(c); return x ? x.v : null; };

  // ---------- leitura do modelo ----------
  const BL_FIELDS = ['Nome da Proposta', 'Cliente', 'Projeto', 'iStudio', 'Papel', 'Senioridade', 'Vaga Intranet', 'Profissional',
    'Custo Profissional', 'R$/H', 'Mês de Execução', 'Horas Provisionadas', 'Horas Executadas', 'Valor Bruto', 'Valor Liquido',
    'Resultado Bruto', 'MB Profissional', 'Status da Receita', 'Obs', 'PIS', 'COFINS', 'INSS', 'ISSQN'];
  const BL_INPUT = ['Nome da Proposta', 'iStudio', 'Papel', 'Senioridade', 'Vaga Intranet', 'Profissional', 'Custo Profissional',
    'R$/H', 'Mês de Execução', 'Horas Provisionadas', 'Horas Executadas', 'Status da Receita', 'Obs'];
  const PROP_FIELDS = ['PowerHouse', 'Proposta', 'Cliente', 'Projeto', 'Ano Fiscal', 'Chance de Conversão', 'iStudio', 'Pre Sales Lead',
    'Start Date', 'End Date', 'COLT OPT', 'Valor Assinado'];

  let uid = 0; const nid = p => p + (++uid).toString(36) + Math.random().toString(36).slice(2, 6);

  async function load(buf, JSZipLib) {
    const JSZ = JSZipLib || root.JSZip;
    const zip = await JSZ.loadAsync(buf);
    const keepBuf = buf;
    const txt = p => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    const wb = await txt('xl/workbook.xml');
    const rels = await txt('xl/_rels/workbook.xml.rels');
    const relMap = {};
    rels.replace(/<Relationship\s([^>]*?)\/>/g, (m, a) => { relMap[attr(a, 'Id')] = attr(a, 'Target'); });
    const sheets = {};
    wb.replace(/<sheet\s([^>]*?)\/>/g, (m, a) => {
      const name = decodeXml(attr(a, 'name')); let t = relMap[attr(a, 'r:id')];
      t = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
      sheets[name] = { path: t, sheetId: +attr(a, 'sheetId') };
    });
    const sst = parseSST(await txt('xl/sharedStrings.xml'));
    // tabelas por aba
    for (const name in sheets) {
      const s = sheets[name]; const rp = s.path.replace(/worksheets\//, 'worksheets/_rels/') + '.rels';
      const r = await txt(rp); s.relsPath = rp; s.tables = [];
      if (r) r.replace(/<Relationship\s([^>]*?)\/>/g, (m, a) => {
        if (/\/table"?$/.test(attr(a, 'Type') || '') || /tables\//.test(attr(a, 'Target'))) {
          const tp = 'xl/' + attr(a, 'Target').replace(/^\.\.\//, ''); s.tables.push(tp);
        }
      });
    }
    const tables = {};
    for (const name in sheets) for (const tp of sheets[name].tables) {
      const x = await txt(tp);
      const tn = attr(x.match(/<table\s[^>]*>/)[0], 'name'); const ref = attr(x.match(/<table\s[^>]*>/)[0], 'ref');
      const cols = []; x.replace(/<tableColumn\s([^>]*?)(\/>|>)/g, (m, a) => { cols.push(decodeXml(attr(a, 'name'))); });
      const trc = +(attr(x.match(/<table\s[^>]*>/)[0], 'totalsRowCount') || 0);
      tables[tn] = { name: tn, path: tp, ref, cols, sheet: name, totals: trc };
      sheets[name].table = tn;
    }
    const parsed = {};
    const getSheet = async name => parsed[name] || (parsed[name] = parseSheet(await txt(sheets[name].path), sst));

    // ---- Config ----
    const cfgName = Object.keys(sheets).find(n => /config/i.test(n));
    const cfg = await getSheet(cfgName);
    const maxRow = Math.max(...cfg.rows.keys());
    const listCol = c => { const out = []; for (let r = 2; r <= maxRow; r++) { const v = cellVal(cfg.rows, r, c); if (v != null && v !== '' && v !== '-' && typeof v === 'string') out.push(v); } return [...new Set(out)]; };
    const lists = {
      statusProposta: listCol(1), revenue: listCol(3), istudio: listCol(5), perfil: listCol(7), senioridade: listCol(9),
      clientes: listCol(15), bp: listCol(16), powerhouse: listCol(18)
    };
    let taxRow = null;
    for (let r = 2; r <= 40; r++) { const u = cellVal(cfg.rows, r, 21); if (typeof u === 'number' && u > 0 && u < 1) { taxRow = r; break; } }
    const taxes = taxRow ? { row: taxRow, PIS: cellVal(cfg.rows, taxRow, 21), COFINS: cellVal(cfg.rows, taxRow, 22), INSS: cellVal(cfg.rows, taxRow, 23), ISSQN: cellVal(cfg.rows, taxRow, 24), OPEX: cellVal(cfg.rows, taxRow, 25) }
      : { row: null, PIS: 0.0065, COFINS: 0.03, INSS: 0.027, ISSQN: 0.02, OPEX: 0.0981 };
    // calendário de dias/horas úteis (Config T:V)
    const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
    const workdays = {};
    for (let r = 2; r <= Math.min(maxRow, 60); r++) {
      const t = cellVal(cfg.rows, r, 20), d = cellVal(cfg.rows, r, 21), h = cellVal(cfg.rows, r, 22);
      const i = typeof t === 'string' ? MESES.indexOf(t.trim().toLowerCase()) : -1;
      if (i >= 0 && typeof d === 'number' && d >= 1 && typeof h === 'number') workdays[i + 1] = { dias: d, horas: h };
    }
    const cfgClients = [];
    for (let r = 2; r <= maxRow; r++) {
      const n = cellVal(cfg.rows, r, 27); if (!n || typeof n !== 'string') continue;
      cfgClients.push({ name: n, sheet: cellVal(cfg.rows, r, 28), table: cellVal(cfg.rows, r, 29), status: cellVal(cfg.rows, r, 31), row: r });
    }
    const keys = [];
    for (let r = 2; r <= maxRow; r++) {
      const k = cellVal(cfg.rows, r, 32); if (!k || typeof k !== 'string') continue;
      keys.push({ key: k, trib: cellVal(cfg.rows, r, 33) || 'Tributado' });
    }

    // ---- clientes (abas BL) ----
    const clients = [];
    const blSheets = Object.keys(sheets).filter(n => /^BL_/.test(n) && n !== 'BL_Template' && sheets[n].table);
    for (const sn of blSheets) {
      const cc = cfgClients.find(c => c.sheet === sn);
      const sh = await getSheet(sn);
      let name = cc ? cc.name : null;
      if (!name) { const t = sh.rows.get(1) && sh.rows.get(1).cells.get(1); const m = t && /—\s*(.+)$/.exec(t.v || ''); name = m ? m[1].trim() : sn.replace(/^BL_/, ''); }
      clients.push({ id: nid('c'), name, sheet: sn, table: sheets[sn].table, status: cc ? cc.status || 'Ativo' : 'Ativo', cfgRow: cc ? cc.row : null });
    }
    const clientByName = n => clients.find(c => norm(c.name) === norm(n));
    const extraClients = [...lists.clientes].filter(n => !clientByName(n));
    for (const n of extraClients) clients.push({ id: nid('c'), name: n, sheet: null, table: null, status: 'Ativo' });

    clients.forEach(c => { c.inList = lists.clientes.some(n => norm(n) === norm(c.name)); });
    // ---- alocações ----
    const allocations = []; const blMeta = {};
    for (const c of clients.filter(c => c.sheet)) {
      const sh = parsed[c.sheet]; const t = tables[c.table];
      const [a, b] = t.ref.split(':').map(splitRef);
      const colOf = {}; t.cols.forEach((h, i) => { colOf[h] = a.c + i; });
      blMeta[c.sheet] = { table: t.name, cols: t.cols, colOf, firstCol: a.c, startRow: a.row + 1, endRow: b.row - (t.totals || 0), totals: t.totals || 0, tplRow: null };
      for (let r = a.row + 1; r <= b.row - (t.totals || 0); r++) {
        const row = sh.rows.get(r); if (!row) continue;
        const g = h => { const x = row.cells.get(colOf[h]); return x ? x.v : null; };
        if (!blMeta[c.sheet].tplRow) blMeta[c.sheet].tplRow = r;
        const proj = g('Projeto'), prof = g('Profissional');
        if ((proj == null || proj === '') && (prof == null || prof === '')) continue;
        const f = {};
        for (const h of t.cols) { let v = g(h); if (h === 'Mês de Execução' && typeof v === 'number') v = serialToISO(v); f[h] = v; }
        const nCell = row.cells.get(colOf['Valor Bruto']);
        let mult = 1; if (nCell && nCell.f) { const mm = /^\s*M\d+\*J\d+\*([\d.]+)\s*$/.exec(nCell.f); if (mm) mult = +mm[1]; }
        let brutoFixed = null; // valor bruto digitado à mão (sem fórmula) e diferente de horas × R$/h
        if (nCell && !nCell.f && typeof nCell.v === 'number' && Math.abs(nCell.v - (+g('Horas Executadas') || 0) * (+g('R$/H') || 0)) > 0.005) brutoFixed = nCell.v;
        const iCell = row.cells.get(colOf['Custo Profissional']);
        const st = row.cells.get(colOf['Status da Receita']);
        if (st && st.f) f['Status da Receita'] = null;
        const extraFormula = t.cols.some(h => !BL_FIELDS.includes(h) && h !== 'Custo Total Profissional' && row.cells.get(colOf[h]) && row.cells.get(colOf[h]).f);
        allocations.push({ id: nid('a'), sheet: c.sheet, project: String(proj == null ? '' : proj), fields: f, mult,
          costFormula: !!(iCell && iCell.f), extraFormula, brutoFixed, orig: { row: r }, dirty: false });
      }
    }

    // ---- propostas (mapeadas pelo nome da coluna) ----
    const propName = Object.keys(sheets).find(n => /propostas/i.test(n));
    const ps = await getSheet(propName); const pt = tables[sheets[propName].table];
    const [pa, pb] = pt.ref.split(':').map(splitRef);
    const pcol = {}; pt.cols.forEach((h, i) => { pcol[h] = pa.c + i; });
    const hasFunding = 'Funding' in pcol;
    const proposals = [];
    for (let r = pa.row + 1; r <= pb.row - (pt.totals || 0); r++) {
      const row = ps.rows.get(r); if (!row) continue;
      const f = {};
      for (const h of [...PROP_FIELDS, 'Funding']) { if (!(h in pcol)) continue; const x = row.cells.get(pcol[h]); let v = x ? x.v : null; if ((h === 'Start Date' || h === 'End Date') && typeof v === 'number') v = serialToISO(v); f[h] = v; }
      if (!f['Cliente'] && !f['Projeto'] && !f['Proposta']) continue;
      const L = row.cells.get(pcol['Valor Assinado']);
      proposals.push({ id: nid('p'), fields: f, lFormula: L && L.f ? L.f : null, orig: { row: r }, dirty: false });
    }

    // ---- projetos (união propostas + alocações + chaves) ----
    const projects = []; const pk = (cl, n) => norm(cl) + '|' + norm(n);
    const findProj = (clientName, name) => projects.find(p => pk(p.client, p.name) === pk(clientName, name));
    const keyTrib = (cl, n) => { const k = keys.find(k => norm(k.key) === pk(cl, n)); return k ? k.trib : null; };
    for (const a of allocations) {
      const c = clients.find(c => c.sheet === a.sheet);
      if (!findProj(c.name, a.project)) projects.push({ id: nid('j'), client: c.name, name: a.project.trim(), trib: null });
    }
    for (const p of proposals) {
      const cl = clientByName(p.fields['Cliente']); const cname = cl ? cl.name : String(p.fields['Cliente'] || '').trim();
      const pn = String(p.fields['Projeto'] || '').trim();
      if (!cl && cname && !clientByName(cname)) clients.push({ id: nid('c'), name: cname, sheet: null, table: null, status: 'Ativo' });
      p.client = cname;
      if (!findProj(cname, pn)) projects.push({ id: nid('j'), client: cname, name: pn, trib: null });
    }
    const model0 = { proposals, hasFunding };
    for (const p of projects) {
      if (hasFunding) continue;
      const c = clientByName(p.client);
      // a chave do Excel usa o Cliente gravado nas linhas BL
      p.trib = keyTrib(p.client, p.name) ||
        (c && c.sheet ? (() => { const a = allocations.find(a => a.sheet === c.sheet && norm(a.project) === norm(p.name)); return a ? keyTrib(a.fields['Cliente'], p.name) : null; })() : null) || 'Tributado';
    }
    if (hasFunding) refreshTrib({ projects, proposals, hasFunding });
    const usedKeys = new Set(projects.map(p => pk(p.client, p.name)));
    const orphanKeys = keys.filter(k => !usedKeys.has(norm(k.key)) && !allocations.some(a => norm(a.fields['Cliente'] + '|' + a.project) === norm(k.key)));

    // ---- pessoas e perfis (aba "Acessos") ----
    const people = []; const accName = Object.keys(sheets).find(n => norm(n) === 'acessos');
    if (accName) {
      const as = await getSheet(accName); const hdr = as.rows.get(1); const hc = {};
      if (hdr) for (const [ci, x] of hdr.cells) if (typeof x.v === 'string') hc[norm(x.v)] = ci;
      const maxR = Math.max(1, ...as.rows.keys());
      for (let r = 2; r <= maxR; r++) {
        const g = k => hc[k] ? cellVal(as.rows, r, hc[k]) : null;
        const nome = g('nome'); if (!nome || typeof nome !== 'string' || !nome.trim()) continue;
        people.push({ id: nid('u'), name: nome.trim(), role: String(g('função') || g('funcao') || 'Delivery Leader').trim(),
          clients: String(g('clientes') || '').split(/\s*;\s*/).map(x => x.trim()).filter(Boolean),
          active: !/^n(ã|a)o$/i.test(String(g('ativo') || 'Sim').trim()) });
      }
    }
    return {
      people, accName,
      zip, sheets, tables, sst, parsedCache: parsed, cfgName, propName, lists, taxes, workdays, cfgClients, orphanKeys,
      clients, projects, proposals, allocations, blMeta, propTable: pt.name, fileName: null, buf: keepBuf, hasFunding,
    };
  }

  // Regra nova: tributação vem do Funding da proposta (Sim = Isento). O Excel usa a 1ª chave encontrada.
  function refreshTrib(m) {
    if (!m.hasFunding) return;
    for (const p of m.projects) {
      const ps = m.proposals.filter(x => norm(x.client || x.fields['Cliente']) === norm(p.client) && norm(x.fields['Projeto']) === norm(p.name));
      const vals = ps.map(x => String(x.fields['Funding'] || '').trim().toLowerCase() === 'sim');
      p.trib = vals.length && vals[0] ? 'Isento' : 'Tributado';
      p.tribConflict = vals.length > 1 && vals.some(v => v !== vals[0]);
      p.noProposal = !ps.length;
    }
  }
  function projectOf(m, clientName, projName) {
    return m.projects.find(p => norm(p.client) === norm(clientName) && norm(p.name) === norm(projName)) || null;
  }
  // ---------- cálculos ----------
  function taxRate(m) { const t = m.taxes; return (t.PIS || 0) + (t.COFINS || 0) + (t.INSS || 0) + (t.ISSQN || 0); }
  function calcAlloc(m, a, projTrib) {
    const f = a.fields; const he = +f['Horas Executadas'] || 0, rate = +f['R$/H'] || 0, custo = +f['Custo Profissional'] || 0;
    const bruto = a.brutoFixed != null ? a.brutoFixed : he * rate * (a.mult || 1);
    const imp = projTrib === 'Isento' ? 0 : bruto * taxRate(m);
    const liq = bruto - imp; const custoTot = custo * he; const res = liq - custoTot;
    return { bruto, imp, liq, custoTot, res, mb: liq ? res / liq : 0 };
  }

  // ---------- exportação ----------
  async function exportXlsx(m, JSZipLib) {
    const JSZ = JSZipLib || root.JSZip;
    const zip = await JSZ.loadAsync(m.buf);
    const sheets = { ...m.sheets }, tables = { ...m.tables }, blMeta = { ...m.blMeta };
    refreshTrib(m);
    const txt = p => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    const sst = m.sst;
    const fresh = async name => parseSheet(await txt(sheets[name].path), sst);
    const clientOf = sheet => m.clients.find(c => c.sheet === sheet);
    const projTrib = (clientName, name) => { const p = m.projects.find(p => norm(p.client) === norm(clientName) && norm(p.name) === norm(name)); return p ? p.trib : 'Tributado'; };
    const taxRow = m.taxes.row || 9;
    const cfgQ = `'${m.cfgName.replace(/'/g, "''")}'`;

    // 1) novas abas BL para clientes novos
    let wbXml = await txt('xl/workbook.xml'); let relsXml = await txt('xl/_rels/workbook.xml.rels');
    let ctXml = await txt('[Content_Types].xml');
    for (const c of m.clients.filter(c => c.isNew && c.createSheet && !sheets[c.sheet])) {
      const tplName = 'BL_Template'; const tpl = sheets[tplName];
      const nSheets = Object.keys(zip.files).filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f)).map(f => +f.match(/(\d+)\.xml$/)[1]);
      const nTables = Object.keys(zip.files).filter(f => /^xl\/tables\/table\d+\.xml$/.test(f)).map(f => +f.match(/(\d+)\.xml$/)[1]);
      const sNum = Math.max(...nSheets) + 1, tNum = Math.max(...nTables) + 1;
      let sx = await txt(tpl.path);
      sx = sx.replace(/\scodeName="[^"]*"/, '').replace(/\stabSelected="1"/, '').replace(/\sxr:uid="[^"]*"/g, '');
      zip.file(`xl/worksheets/sheet${sNum}.xml`, sx);
      let tx = await txt(tpl.tables[0]);
      const ids = []; for (const tp of Object.values(tables)) { const x = await txt(tp.path); ids.push(+attr(x.match(/<table\s[^>]*>/)[0], 'id')); }
      const newId = Math.max(...ids) + 1;
      tx = tx.replace(/(<table\s[^>]*?)\sid="\d+"/, `$1 id="${newId}"`).replace(/name="tbl_Template"/, `name="${c.table}"`)
        .replace(/displayName="tbl_Template"/, `displayName="${c.table}"`).replace(/\sxr3?:uid="[^"]*"/g, '');
      zip.file(`xl/tables/table${tNum}.xml`, tx);
      zip.file(`xl/worksheets/_rels/sheet${sNum}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tNum}.xml"/></Relationships>`);
      const rid = 'rId' + (Math.max(...[...relsXml.matchAll(/Id="rId(\d+)"/g)].map(x => +x[1])) + 1);
      relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sNum}.xml"/></Relationships>`);
      const sid = Math.max(...[...wbXml.matchAll(/sheetId="(\d+)"/g)].map(x => +x[1])) + 1;
      wbXml = wbXml.replace('</sheets>', `<sheet name="${encXml(c.sheet)}" sheetId="${sid}" r:id="${rid}"/></sheets>`);
      ctXml = ctXml.replace('</Types>', `<Override PartName="/xl/worksheets/sheet${sNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table${tNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`);
      sheets[c.sheet] = { path: `xl/worksheets/sheet${sNum}.xml`, tables: [`xl/tables/table${tNum}.xml`], relsPath: `xl/worksheets/_rels/sheet${sNum}.xml.rels` };
      tables[c.table] = { name: c.table, path: `xl/tables/table${tNum}.xml`, ref: tables['tbl_Template'].ref, cols: tables['tbl_Template'].cols, sheet: c.sheet };
      const tm = tables['tbl_Template']; const [a] = tm.ref.split(':').map(splitRef);
      const colOf = {}; tm.cols.forEach((h, i) => { colOf[h] = a.c + i; });
      blMeta[c.sheet] = { table: c.table, cols: tm.cols, colOf, firstCol: a.c, startRow: a.row + 1, endRow: a.row + 1, tplRow: a.row + 1, fromTemplate: true };
    }
    // remove cadeia de cálculo (as fórmulas mudaram) e força recálculo ao abrir
    if (zip.file('xl/calcChain.xml')) {
      zip.remove('xl/calcChain.xml');
      relsXml = relsXml.replace(/<Relationship\s[^>]*calcChain[^>]*\/>/, '');
      ctXml = ctXml.replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, '');
    }
    wbXml = /<calcPr\s/.test(wbXml) ? wbXml.replace(/<calcPr\s([^>]*?)\/>/, (x, a) => `<calcPr ${a.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`)
      : wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/workbook.xml', wbXml); zip.file('xl/_rels/workbook.xml.rels', relsXml); zip.file('[Content_Types].xml', ctXml);

    // 2) abas BL
    for (const c of m.clients.filter(c => c.sheet && blMeta[c.sheet])) {
      const meta = blMeta[c.sheet]; const sh = await fresh(c.sheet);
      const L = h => numToCol(meta.colOf[h]);
      const recs = m.allocations.filter(a => a.sheet === c.sheet);
      const tplRow = sh.rows.get(meta.tplRow) || sh.rows.get(meta.startRow);
      const tplStyles = new Map(); if (tplRow) for (const [ci, x] of tplRow.cells) tplStyles.set(ci, x.s);
      if (meta.fromTemplate) { // o BL_Template pode ter formatos incompletos: copia a formatação de uma aba BL existente
        const ref = Object.entries(blMeta).find(([sn, mm]) => !mm.fromTemplate && sheets[sn] && mm.cols.join('|') === meta.cols.join('|'));
        if (ref) { const rs = await fresh(ref[0]); const rr = rs.rows.get(ref[1].tplRow || ref[1].startRow);
          if (rr) meta.cols.forEach((h, i) => { const x = rr.cells.get(ref[1].colOf[h]); if (x && x.s != null) tplStyles.set(meta.firstCol + i, x.s); }); }
      }
      const tplAttrs = (tplRow ? tplRow.attrs : `r="${meta.startRow}"`).replace(/\s*hidden="1"/, '');
      const newRows = new Map();
      for (const [r, row] of sh.rows) if (r < meta.startRow) newRows.set(r, row);
      let r = meta.startRow;
      const lastC = meta.firstCol + meta.cols.length - 1;
      const emit = (rec) => {
        const proj = rec ? projectOf(m, c.name, rec.project) : null;
        const calc = rec ? calcAlloc(m, rec, proj ? proj.trib : 'Tributado') : null;
        const orow = rec && rec.orig && rec.orig.row ? sh.rows.get(rec.orig.row) : null;
        const dr = orow ? r - rec.orig.row : 0;
        const cells = new Map();
        for (let ci = meta.firstCol; ci <= lastC; ci++) {
          const h = meta.cols[ci - meta.firstCol]; const ref = numToCol(ci) + r;
          const oc = orow ? orow.cells.get(ci) : null; const s = oc ? oc.s : tplStyles.get(ci);
          let xml;
          if (!rec) { xml = canonical(h, ref, s, oc, dr) || `<c r="${ref}"${s != null ? ` s="${s}"` : ''}/>`; }
          else {
            const canon = canonical(h, ref, s, oc, dr, rec, calc);
            if (canon) xml = canon;
            else if (h === 'Cliente') xml = cellXml(ref, s, c.name);
            else if (h === 'Projeto') xml = cellXml(ref, s, proj ? proj.name : rec.project);
            else if (h === 'Custo Profissional' && rec.costFormula && oc && oc.f && !rec.costEdited) xml = keepCellXml(oc, ref, dr);
            else if (oc && oc.f && !rec.dirty && !BL_INPUT.includes(h)) xml = keepCellXml(oc, ref, dr);
            else { const v = rec.fields[h]; xml = cellXml(ref, s, v, { date: h === 'Mês de Execução' }); }
          }
          cells.set(ci, { xml });
        }
        // células fora da tabela na mesma linha original (raro): preserva
        if (orow) for (const [ci, x] of orow.cells) if (ci < meta.firstCol || ci > lastC) cells.set(ci, { xml: keepCellXml(x, numToCol(ci) + r, dr) });
        const attrs = (orow ? orow.attrs : tplAttrs).replace(/\s*hidden="1"/, '');
        newRows.set(r, { attrs, cells }); r++;
      };
      function canonical(h, ref, s, oc, dr, rec, k) {
        const R = ref.replace(/^[A-Z]+/, '');
        const sa = s;
        const isen = k && k.imp === 0 && k.bruto !== 0;
        const tv = name => k ? (isen ? 0 : (m.taxes[name] || 0) * k.bruto) : undefined;
        const vals = k ? { 'Valor Bruto': k.bruto, 'Valor Liquido': k.liq, 'Resultado Bruto': k.res, 'MB Profissional': k.mb,
          PIS: tv('PIS'), COFINS: tv('COFINS'), INSS: tv('INSS'), ISSQN: tv('ISSQN'), 'Custo Total Profissional': k.custoTot } : {};
        const F = f => cellXml(ref, sa, null, { f, v: vals[h] });
        const T = (col) => `IF(IFERROR(MATCH(${L('Cliente')}${R}&"|"&${L('Projeto')}${R},${cfgQ}!$AF$2:$AF$400,0),0)>0,IF(INDEX(${cfgQ}!$AG$2:$AG$400,MATCH(${L('Cliente')}${R}&"|"&${L('Projeto')}${R},${cfgQ}!$AF$2:$AF$400,0))="Isento",0,${cfgQ}!${col}$${taxRow}*${L('Valor Bruto')}${R}),${cfgQ}!${col}$${taxRow}*${L('Valor Bruto')}${R})`;
        switch (h) {
          case 'Valor Bruto': {
            if (rec && rec.brutoFixed != null) return cellXml(ref, sa, rec.brutoFixed);
            const mult = rec ? rec.mult : 1;
            return F(`${L('Horas Executadas')}${R}*${L('R$/H')}${R}` + (mult && mult !== 1 ? `*${mult}` : ''));
          }
          case 'Valor Liquido': return F(`${L('Valor Bruto')}${R}-(${L('PIS')}${R}+${L('COFINS')}${R}+${L('INSS')}${R}+${L('ISSQN')}${R})`);
          case 'Resultado Bruto': return F(`${L('Valor Liquido')}${R}-(${L('Custo Profissional')}${R}*${L('Horas Executadas')}${R})`);
          case 'MB Profissional': return F(`IF(${L('Valor Liquido')}${R}=0,0,${L('Resultado Bruto')}${R}/${L('Valor Liquido')}${R})`);
          case 'PIS': return F(T('U')); case 'COFINS': return F(T('V')); case 'INSS': return F(T('W')); case 'ISSQN': return F(T('X'));
          case 'Custo Total Profissional': return F(`${meta.table}[[#This Row],[Custo Profissional]]*${meta.table}[[#This Row],[Horas Executadas]]`);
          default:
            if (!BL_FIELDS.includes(h) && rec && oc && oc.f) return keepCellXml(oc, ref, dr);
            if (!BL_FIELDS.includes(h) && rec) return cellXml(ref, s, rec.fields[h]);
            return null;
        }
      }
      for (const rec of recs) emit(rec);
      if (!recs.length) emit(null);
      const dataEnd = r - 1;
      for (let k = 1; k <= (meta.totals || 0); k++) { // linha de totais da tabela
        const orow = sh.rows.get(meta.endRow + k); const cells = new Map();
        if (orow) for (const [ci, x] of orow.cells) cells.set(ci, { xml: keepCellXml(x, numToCol(ci) + r, r - (meta.endRow + k)) });
        newRows.set(r, { attrs: orow ? orow.attrs : `r="${r}"`, cells }); r++;
      }
      const endRow = r - 1;
      const lastCol = numToCol(Math.max(lastC, ...[...newRows.values()].flatMap(x => [...x.cells.keys()])));
      const xml = sh.head.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${lastCol}${endRow}"/>`) + '<sheetData>' + serializeRows(newRows) + '</sheetData>' + sh.tail;
      zip.file(sheets[c.sheet].path, xml);
      // tabela
      const tp = tables[meta.table].path; let tx = await txt(tp);
      const ref = `${numToCol(meta.firstCol)}${meta.startRow - 1}:${numToCol(lastC)}${endRow}`;
      tx = tx.replace(/(<table\s[^>]*?\sref=")[^"]*"/, `$1${ref}"`)
        .replace(/<autoFilter\s([^>]*?)(\/>|>[\s\S]*?<\/autoFilter>)/, (x, a) => `<autoFilter ${a.replace(/ref="[^"]*"/, `ref="${numToCol(meta.firstCol)}${meta.startRow - 1}:${numToCol(lastC)}${dataEnd}"`)}/>`)
        .replace(/<sortState[\s\S]*?<\/sortState>/, '');
      zip.file(tp, tx);
      if (c.isNew) { // título da nova aba
        const sx = await txt(sheets[c.sheet].path);
        zip.file(sheets[c.sheet].path, sx.replace(/(<c r="A1"[^>]*?)(\/>|>[\s\S]*?<\/c>)/, (x, a) => cellXml('A1', attr(a, 's'), 'Billing — ' + c.name)));
      }
    }

    // 3) Propostas (colunas pelo nome)
    const propRowOf = []; // linha gravada de cada proposta (para as chaves da Config)
    {
      refreshTrib(m);
      const sh = await fresh(m.propName); const t = tables[m.propTable];
      const [a, b] = t.ref.split(':').map(splitRef);
      const col = {}; t.cols.forEach((h, i) => { col[h] = a.c + i; });
      const lastC = a.c + t.cols.length - 1; const CL = h => numToCol(col[h]);
      const tplRow = sh.rows.get(a.row + 1);
      const tplStyles = new Map(); if (tplRow) for (const [ci, x] of tplRow.cells) tplStyles.set(ci, x.s);
      const newRows = new Map(); for (const [r, row] of sh.rows) if (r <= a.row) newRows.set(r, row);
      const after = [...sh.rows.entries()].filter(([r]) => r > b.row);
      let r = a.row + 1;
      const allTables = m.clients.filter(c => c.table).map(c => c.table);
      for (const p of m.proposals) {
        const orow = p.orig && p.orig.row ? sh.rows.get(p.orig.row) : null; const dr = orow ? r - p.orig.row : 0;
        const cells = new Map();
        const cl = m.clients.find(c => norm(c.name) === norm(p.client));
        const pj = projectOf(m, p.client, p.fields['Projeto']);
        const projName = pj ? pj.name : p.fields['Projeto'];
        const usado = cl && cl.sheet ? m.allocations.filter(x => x.sheet === cl.sheet && norm(x.project) === norm(p.fields['Projeto'])).reduce((tt, x) => tt + calcAlloc(m, x, 'Tributado').bruto, 0) : 0;
        const Lv = typeof p.fields['Valor Assinado'] === 'number' ? p.fields['Valor Assinado'] : 0;
        const clientName = cl ? cl.name : (p.client || p.fields['Cliente']);
        for (let ci = a.c; ci <= lastC; ci++) {
          const h = t.cols[ci - a.c]; const ref = numToCol(ci) + r; const oc = orow ? orow.cells.get(ci) : null; const s = oc ? oc.s : tplStyles.get(ci);
          let xml;
          if (h === 'Valor Assinado' && p.lFormula && !p.lEdited && oc && oc.f) xml = keepCellXml(oc, ref, dr);
          else if (h === 'Cliente') xml = cellXml(ref, s, clientName);
          else if (h === 'Projeto') xml = cellXml(ref, s, projName);
          else if (h === 'Projeto - Config') xml = cellXml(ref, s, null, { f: `_xlfn.CONCAT(${t.name}[[#This Row],[Cliente]],"|",${t.name}[[#This Row],[Projeto]])`, v: `${clientName}|${projName}` });
          else if (h === 'Valor Usado') {
            const E = CL('Projeto') + r;
            const f = cl && cl.table ? `SUMIF(${cl.table}[Projeto],TRIM(${E}),${cl.table}[Valor Bruto])` : allTables.map(tb => `SUMIF(${tb}[Projeto],TRIM(${E}),${tb}[Valor Bruto])`).join('+');
            xml = f ? cellXml(ref, s, null, { f, v: usado }) : cellXml(ref, s, 0); // sem nenhuma aba BL ainda: valor 0 (fórmula vazia corrompe o arquivo)
          } else if (h === 'Delta') xml = cellXml(ref, s, null, { f: `${CL('Valor Assinado')}${r}-${CL('Valor Usado')}${r}`, v: Lv - usado });
          else if (PROP_FIELDS.includes(h) || h === 'Funding') xml = cellXml(ref, s, p.fields[h], { date: h === 'Start Date' || h === 'End Date' });
          else xml = oc ? keepCellXml(oc, ref, dr) : cellXml(ref, s, null);
          cells.set(ci, { xml });
        }
        propRowOf.push({ p, r, key: `${clientName}|${projName}`, isento: String(p.fields['Funding'] || '').trim().toLowerCase() === 'sim' });
        newRows.set(r, { attrs: (orow ? orow.attrs : (tplRow ? tplRow.attrs : `r="${r}"`)).replace(/\s*hidden="1"/, ''), cells }); r++;
      }
      const endRow = Math.max(r - 1, a.row + 1);
      const shiftAfter = endRow - b.row;
      for (const [or, row] of after) {
        const nr = or + shiftAfter; const cells = new Map();
        for (const [ci, x] of row.cells) cells.set(ci, { xml: keepCellXml(x, numToCol(ci) + nr, shiftAfter) });
        newRows.set(nr, { attrs: row.attrs, cells });
      }
      const maxR = Math.max(...newRows.keys());
      const xml = sh.head.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${numToCol(lastC)}${maxR}"/>`) + '<sheetData>' + serializeRows(newRows) + '</sheetData>' + sh.tail;
      zip.file(sheets[m.propName].path, xml);
      const ref = `${numToCol(a.c)}${a.row}:${numToCol(lastC)}${endRow}`;
      let tx = await txt(t.path);
      tx = tx.replace(/(<table\s[^>]*?\sref=")[^"]*"/, `$1${ref}"`)
        .replace(/<autoFilter\s([^>]*?)(\/>|>[\s\S]*?<\/autoFilter>)/, (x, aa) => `<autoFilter ${aa.replace(/ref="[^"]*"/, `ref="${ref}"`)}/>`)
        .replace(/<sortState[\s\S]*?<\/sortState>/, '');
      zip.file(t.path, tx);
      m._propCols = { key: col['Projeto - Config'] ? numToCol(col['Projeto - Config']) : null, funding: col['Funding'] ? numToCol(col['Funding']) : null };
    }

    // 4) Config: chaves AF/AG, clientes AA:AE, lista O
    {
      const sh = await fresh(m.cfgName); const rows = sh.rows;
      const setCell = (r, c, v) => {
        if (!rows.has(r)) rows.set(r, { attrs: `r="${r}"`, cells: new Map() });
        const row = rows.get(r); const old = row.cells.get(c); const ref = numToCol(c) + r;
        row.cells.set(c, { xml: cellXml(ref, old ? old.s : null, v) });
      };
      const propQ = `'${m.propName.replace(/'/g, "''")}'`;
      const useFormulas = m.hasFunding && m._propCols && m._propCols.key && m._propCols.funding;
      // chaves: usa o nome do cliente gravado nas linhas BL (é o que a fórmula compara)
      const entries = []; const seen = new Set();
      if (useFormulas) {
        // regra nova: cada linha da Config aponta para uma linha de Propostas (chave + Funding)
        const setF = (r, c, f, v) => { if (!rows.has(r)) rows.set(r, { attrs: `r="${r}"`, cells: new Map() }); const row = rows.get(r); const old = row.cells.get(c); row.cells.set(c, { xml: cellXml(numToCol(c) + r, old ? old.s : null, null, { f, v }) }); };
        const maxR2 = Math.max(...rows.keys(), 102, propRowOf.length + 1);
        for (let r = 2; r <= maxR2; r++) {
          const e = propRowOf[r - 2];
          if (e) { setF(r, 32, `${propQ}!${m._propCols.key}${e.r}`, e.key); setF(r, 33, `IF(${propQ}!${m._propCols.funding}${e.r}="Sim","Isento","Tributado")`, e.isento ? 'Isento' : 'Tributado'); }
          else if (rows.has(r)) { const row = rows.get(r); row.cells.delete(32); row.cells.delete(33); }
        }
      }
      if (!useFormulas) {
      for (const p of m.projects) {
        const cl = m.clients.find(c => norm(c.name) === norm(p.client));
        const k = `${cl ? cl.name : p.client}|${p.name}`;
        if (!seen.has(norm(k))) { seen.add(norm(k)); entries.push([k, p.trib || 'Tributado']); }
      }
      for (const k of m.orphanKeys) if (!seen.has(norm(k.key))) { seen.add(norm(k.key)); entries.push([k.key, k.trib]); }
      const maxR = Math.max(...rows.keys(), 102, entries.length + 1);
      for (let r = 2; r <= maxR; r++) {
        if (r - 2 < entries.length) { setCell(r, 32, entries[r - 2][0]); setCell(r, 33, entries[r - 2][1]); }
        else if (rows.has(r)) { const row = rows.get(r); row.cells.delete(32); row.cells.delete(33); }
      }
      }
      // espelha no Database as alíquotas vindas do Config.xlsx (as fórmulas BL leem daqui)
      if (m.config && m.config.taxesFromConfig && m.taxes.row) { const tr = m.taxes.row; [['PIS', 21], ['COFINS', 22], ['INSS', 23], ['ISSQN', 24], ['OPEX', 25]].forEach(([k, c]) => { if (m.taxes[k] != null) setCell(tr, c, m.taxes[k]); }); }
      // clientes novos
      let lastAA = Math.max(1, ...[...rows.entries()].filter(([, x]) => x.cells.get(27) && x.cells.get(27).v).map(([r]) => r));
      let lastO = Math.max(1, ...[...rows.entries()].filter(([, x]) => x.cells.get(15) && x.cells.get(15).v).map(([r]) => r));
      for (const c of m.clients.filter(c => c.isNew)) {
        if (c.sheet && !c.cfgRow) { lastAA++; setCell(lastAA, 27, c.name); setCell(lastAA, 28, c.sheet); setCell(lastAA, 29, c.table); setCell(lastAA, 31, c.status || 'Ativo'); }
        if (!c.inList) { lastO++; setCell(lastO, 15, c.name); }
      }
      let head = sh.head; const maxRow = Math.max(...rows.keys());
      head = head.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:AG${maxRow}"/>`);
      zip.file(sheets[m.cfgName].path, head + '<sheetData>' + serializeRows(rows) + '</sheetData>' + sh.tail);
    }

    // 5) Aba "Acessos" (pessoas, função e clientes)
    if (m.people && (m.people.length || m.accName)) {
      const hdrs = ['Nome', 'Função', 'Clientes', 'Ativo'];
      let rowsXml = `<row r="1">${hdrs.map((h, i) => cellXml(numToCol(i + 1) + '1', null, h)).join('')}</row>`;
      m.people.forEach((p, i) => { const r = i + 2; rowsXml += `<row r="${r}">` + [p.name, p.role, (p.clients || []).join('; '), p.active === false ? 'Não' : 'Sim'].map((v, j) => cellXml(numToCol(j + 1) + r, null, v)).join('') + '</row>'; });
      const last = m.people.length + 1;
      const body = `<dimension ref="A1:D${last}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="3" width="70" customWidth="1"/><col min="4" max="4" width="8" customWidth="1"/></cols><sheetData>${rowsXml}</sheetData>`;
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</worksheet>`;
      if (m.accName && sheets[m.accName]) zip.file(sheets[m.accName].path, xml);
      else {
        let wbX = await txt('xl/workbook.xml'), relX = await txt('xl/_rels/workbook.xml.rels'), ctX = await txt('[Content_Types].xml');
        const nS = Math.max(...Object.keys(zip.files).filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f)).map(f => +f.match(/(\d+)\.xml$/)[1])) + 1;
        zip.file(`xl/worksheets/sheet${nS}.xml`, xml);
        const rid = 'rId' + (Math.max(...[...relX.matchAll(/Id="rId(\d+)"/g)].map(x => +x[1])) + 1);
        relX = relX.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${nS}.xml"/></Relationships>`);
        const sid = Math.max(...[...wbX.matchAll(/sheetId="(\d+)"/g)].map(x => +x[1])) + 1;
        wbX = wbX.replace('</sheets>', `<sheet name="Acessos" sheetId="${sid}" r:id="${rid}"/></sheets>`);
        ctX = ctX.replace('</Types>', `<Override PartName="/xl/worksheets/sheet${nS}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
        zip.file('xl/workbook.xml', wbX); zip.file('xl/_rels/workbook.xml.rels', relX); zip.file('[Content_Types].xml', ctX);
      }
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }


  // ================= Config.xlsx (regras) =================
  // Lê as abas do Config e sobrepõe as regras do modelo carregado do Database.
  async function loadConfig(m, buf, JSZipLib) {
    const JSZ = JSZipLib || root.JSZip;
    const zip = await JSZ.loadAsync(buf);
    const txt = p => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    const wb = await txt('xl/workbook.xml'); const rels = await txt('xl/_rels/workbook.xml.rels');
    const relMap = {}; rels.replace(/<Relationship\s([^>]*?)\/>/g, (x, a) => { relMap[attr(a, 'Id')] = attr(a, 'Target'); });
    const sheets = {}; wb.replace(/<sheet\s([^>]*?)\/>/g, (x, a) => { const name = decodeXml(attr(a, 'name')); let t = relMap[attr(a, 'r:id')]; t = t.startsWith('/') ? t.slice(1) : 'xl/' + t; sheets[name] = { path: t }; });
    const sst = parseSST(await txt('xl/sharedStrings.xml'));
    const find = re => Object.keys(sheets).find(n => re.test(norm(n)));
    const rowsOf = async name => { if (!name) return null; return parseSheet(await txt(sheets[name].path), sst).rows; };
    const table = async name => { // linhas como objetos pelo cabeçalho da linha 1
      const rows = await rowsOf(name); if (!rows) return [];
      const hdr = rows.get(1); if (!hdr) return []; const hc = []; for (const [ci, x] of hdr.cells) if (x.v != null && x.v !== '') hc.push([ci, String(x.v).trim()]);
      const out = []; const maxR = Math.max(1, ...rows.keys());
      for (let r = 2; r <= maxR; r++) { const o = { _row: r }; let any = false; for (const [ci, h] of hc) { const v = cellVal(rows, r, ci); o[h] = v; if (v != null && v !== '') any = true; } if (any) out.push(o); }
      return out;
    };
    const col1 = async name => (await table(name)).map(o => Object.values(o).find((v, i) => i > 0 && v != null)).filter(v => typeof v === 'string' && v.trim() && v !== '-');
    const names = { status: find(/^status da proposta$/), revenue: find(/^revenue state$/), perfis: find(/^perfis$/), periodo: find(/^periodo$/), bp: find(/^bp por cliente$/),
      projetos: find(/^projetos$/), clientes: find(/^clientes x tabelas$/), dias: find(/^dias uteis$/), ph: find(/^powerhouse$/), impostos: find(/^impostos$/) };
    const C = { zip, sheets, sst, names, buf };
    // listas
    if (names.status) m.lists.statusProposta = await col1(names.status);
    if (names.revenue) m.lists.revenue = await col1(names.revenue);
    if (names.ph) m.lists.powerhouse = await col1(names.ph);
    // perfis / ratecard
    if (names.perfis) {
      const P = await table(names.perfis); const key = (o, re) => Object.keys(o).find(k => re.test(norm(k)));
      C.perfis = P.map(o => ({ istudio: String(o[key(o, /studio/)] || '').trim(), perfil: String(o[key(o, /^perfil$/)] || '').trim(), senioridade: String(o[key(o, /senioridade/)] || '').trim(),
        custo2026: +o[key(o, /custo total 2026/)] || null, rate2026: +o[key(o, /ratecard.*2026/)] || null, custo2027: +o[key(o, /2027 pc/)] || null, rate2027: +o[key(o, /ratecard.*2027/)] || null })).filter(p => p.perfil);
      const uniq = a => [...new Set(a.filter(Boolean))].sort((x, y) => x.localeCompare(y));
      // Config.xlsx é a fonte única: as listas da aba "1. Config" da Database deixam de ser somadas
      if (C.perfis.length) {
        const SEN = ['trainee', 'junior', 'pleno', 'senior', 'especialista']; const sk = s => SEN.indexOf(norm(s).normalize('NFD').replace(/[\u0300-\u036f]/g, ''));
        m.lists.perfil = uniq(C.perfis.map(p => p.perfil));
        m.lists.istudio = uniq(C.perfis.map(p => p.istudio));
        m.lists.senioridade = [...new Set(C.perfis.map(p => p.senioridade).filter(Boolean))].sort((a, b) => (sk(a) < 0 ? 99 : sk(a)) - (sk(b) < 0 ? 99 : sk(b)) || a.localeCompare(b));
      }
    }
    // impostos
    if (names.impostos) { const T = (await table(names.impostos))[0]; if (T) { const g = re => { const k = Object.keys(T).find(k => re.test(norm(k))); return k ? +T[k] : null; };
      const t = { PIS: g(/^pis$/), COFINS: g(/^cofins$/), INSS: g(/^inss$/), ISSQN: g(/^issqn$/), OPEX: g(/opex/) };
      for (const k in t) if (t[k] != null && isFinite(t[k])) m.taxes[k] = t[k]; C.taxesFromConfig = true; } }
    // dias úteis
    if (names.dias) { const MES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
      for (const o of await table(names.dias)) { const v = Object.values(o).slice(1); const i = MES.indexOf(norm(v[0])); if (i >= 0 && +v[1] > 0) m.workdays[i + 1] = { dias: +v[1], horas: +v[2] }; } }
    // BP por cliente e status dos projetos (preservados ao regravar)
    if (names.bp) { C.bp = (await table(names.bp)).map(o => { const v = Object.values(o).slice(1); return { cliente: v[0], bp: v[1] }; }).filter(x => x.cliente && x.cliente !== '-'); m.lists.bp = [...new Set(C.bp.map(x => x.bp).filter(Boolean))]; }
    if (names.projetos) { C.projStatus = {}; for (const o of await table(names.projetos)) { const v = Object.values(o).slice(1); if (v[1]) C.projStatus[norm(v[1])] = v[0]; } }
    // clientes da Config que ainda não existem no Database (ficam disponíveis para novos projetos)
    if (names.clientes) {
      const cfgCli = [];
      for (const o of await table(names.clientes)) { const v = Object.values(o).slice(1); const n = typeof v[0] === 'string' ? v[0].trim() : null; if (n) cfgCli.push(n);
        if (n && !m.clients.some(c => norm(c.name) === norm(n))) m.clients.push({ id: nid('c'), name: n, sheet: null, table: null, status: 'Ativo', inList: true }); }
      // clientes que vinham só da lista da Database (sem aba BL_, fora do Config e sem projeto/contrato) saem dos formulários
      const inCfg = n => cfgCli.some(x => norm(x) === norm(n));
      const used = n => m.projects.some(p => norm(p.client) === norm(n)) || m.proposals.some(p => norm(p.client || p.fields['Cliente']) === norm(n));
      m.clients = m.clients.filter(c => c.sheet || c.isNew || inCfg(c.name) || used(c.name));
      m.lists.clientes = m.clients.map(c => c.name);
    }
    m.config = C;
    return m;
  }
  async function exportConfig(m, JSZipLib) {
    const C = m.config; if (!C) return null;
    const JSZ = JSZipLib || root.JSZip; const zip = await JSZ.loadAsync(C.buf);
    const txt = p => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    refreshTrib(m);
    const rewrite = async (name, rowsData, ncols) => { // reescreve a partir da linha 2, mantendo o cabeçalho e o estilo da linha 2
      if (!name) return; const sh = parseSheet(await txt(C.sheets[name].path), C.sst);
      const tpl = sh.rows.get(2); const st = c => tpl && tpl.cells.get(c) ? tpl.cells.get(c).s : null;
      const rows = new Map(); for (const [r, row] of sh.rows) if (r < 2) rows.set(r, row);
      rowsData.forEach((vals, i) => { const r = i + 2; const cells = new Map(); for (let c = 1; c <= ncols; c++) cells.set(c, { xml: cellXml(numToCol(c) + r, st(c), vals[c - 1]) }); rows.set(r, { attrs: tpl ? tpl.attrs.replace(/\s*hidden="1"/, '') : `r="${r}"`, cells }); });
      for (const [r, row] of sh.rows) if (r >= 2 + rowsData.length) { const cells = new Map(); for (const [ci, x] of row.cells) if (ci > ncols) cells.set(ci, { xml: keepCellXml(x, numToCol(ci) + r, 0) }); if (cells.size) rows.set(r, { attrs: row.attrs, cells }); }
      const last = Math.max(1, ...rows.keys());
      zip.file(C.sheets[name].path, sh.head.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${numToCol(ncols)}${last}"/>`) + '<sheetData>' + serializeRows(rows) + '</sheetData>' + sh.tail);
    };
    // Projetos: uma linha por contrato (chave e tributação como valores, sem vínculo com outro arquivo)
    const seen = new Set(); const proj = [];
    for (const p of m.proposals) {
      const cl = m.clients.find(c => norm(c.name) === norm(p.client)); const pj = projectOf(m, p.client, p.fields['Projeto']);
      const key = `${cl ? cl.name : p.client}|${pj ? pj.name : p.fields['Projeto']}`;
      const trib = m.hasFunding ? (String(p.fields['Funding'] || '').trim().toLowerCase() === 'sim' ? 'Isento' : 'Tributado') : (pj ? pj.trib : 'Tributado');
      proj.push([(C.projStatus && C.projStatus[norm(key)]) || 'Ativo', key, trib]); seen.add(norm(key));
    }
    await rewrite(C.names.projetos, proj, 3);
    // Clientes x Tabelas: todos os clientes com aba BL no Database
    await rewrite(C.names.clientes, m.clients.filter(c => c.sheet).map(c => [c.name, c.sheet, c.table]), 3);
    // BP por Cliente: mantém os existentes e acrescenta clientes novos
    if (C.names.bp) { const list = (C.bp || []).map(x => [x.cliente, x.bp || null]); for (const c of m.clients) if (!list.some(x => norm(x[0]) === norm(c.name))) list.push([c.name, null]); await rewrite(C.names.bp, [['-', null], ...list], 2); }
    // remove o vínculo externo com a planilha antiga
    let wbX = await txt('xl/workbook.xml'), relX = await txt('xl/_rels/workbook.xml.rels'), ctX = await txt('[Content_Types].xml');
    if (/externalReference/.test(wbX)) {
      wbX = wbX.replace(/<externalReferences>[\s\S]*?<\/externalReferences>/, '').replace(/<definedName[^>]*>[^<]*\[\d+\][^<]*<\/definedName>/g, '');
      relX = relX.replace(/<Relationship\s[^>]*externalLink[^>]*\/>/g, '');
      ctX = ctX.replace(/<Override PartName="\/xl\/externalLinks\/[^"]*"[^>]*\/>/g, '');
      Object.keys(zip.files).filter(f => /^xl\/externalLinks\//.test(f)).forEach(f => zip.remove(f));
    }
    if (zip.file('xl/calcChain.xml')) { zip.remove('xl/calcChain.xml'); relX = relX.replace(/<Relationship\s[^>]*calcChain[^>]*\/>/, ''); ctX = ctX.replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, ''); }
    wbX = /<calcPr\s/.test(wbX) ? wbX.replace(/<calcPr\s([^>]*?)\/>/, (x, a) => `<calcPr ${a.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`) : wbX;
    zip.file('xl/workbook.xml', wbX); zip.file('xl/_rels/workbook.xml.rels', relX); zip.file('[Content_Types].xml', ctX);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  // ================= DRP (somente leitura) =================
  // Lê a aba "Custo Referência" da planilha DRP fornecida pela gestão: Innovation Studios, perfis,
  // senioridades e, por ano, salário, personnel cost, custo com/sem overhead e ratecard.
  const plainTxt = s => norm(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  async function loadDRP(buf, JSZipLib) {
    const JSZ = JSZipLib || root.JSZip;
    const zip = await JSZ.loadAsync(buf);
    const txt = p => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    const wb = await txt('xl/workbook.xml'); const rels = await txt('xl/_rels/workbook.xml.rels');
    if (!wb || !rels) throw new Error('o arquivo não é uma planilha Excel válida');
    const relMap = {}; rels.replace(/<Relationship\s([^>]*?)\/>/g, (x, a) => { relMap[attr(a, 'Id')] = attr(a, 'Target'); });
    let sheetName = null, sheetPath = null;
    wb.replace(/<sheet\s([^>]*?)\/>/g, (x, a) => {
      const name = decodeXml(attr(a, 'name'));
      if (!sheetPath && /^custo referencia$/.test(plainTxt(name))) { let t = relMap[attr(a, 'r:id')]; sheetPath = t.startsWith('/') ? t.slice(1) : 'xl/' + t; sheetName = name; }
    });
    if (!sheetPath) throw new Error('não encontrei a aba "Custo Referência"');
    const sh = parseSheet(await txt(sheetPath), parseSST(await txt('xl/sharedStrings.xml')));
    // linha de cabeçalho: a primeira (até a 15ª) que tem "INNOVATION STUDIO"
    let hr = null; const head = {};
    for (let r = 1; r <= 15 && hr == null; r++) {
      const row = sh.rows.get(r); if (!row) continue;
      for (const [c, x] of row.cells) if (typeof x.v === 'string' && /innovation studio/.test(plainTxt(x.v))) hr = r;
    }
    if (hr == null) throw new Error('a aba "Custo Referência" não tem a coluna INNOVATION STUDIO');
    for (const [c, x] of sh.rows.get(hr).cells) if (typeof x.v === 'string' && x.v.trim()) head[c] = plainTxt(x.v);
    const colOf = re => { for (const c in head) if (re.test(head[c])) return +c; return null; };
    const cols = { contrato: colOf(/^contrato$/), istudio: colOf(/innovation studio/), perfil: colOf(/^perfil$/), senioridade: colOf(/^senioridade$/), completo: colOf(/^perfil completo$/) };
    if (!cols.perfil || !cols.senioridade) throw new Error('a aba "Custo Referência" precisa das colunas Perfil e Senioridade');
    const yearCols = {}; // ano -> { salario, pc, custoOH, custo, rate }
    const setY = (y, k, c) => { (yearCols[y] = yearCols[y] || {})[k] = +c; };
    for (const c in head) {
      const h = head[c]; let m;
      if ((m = /^salario referencia (\d{4})/.exec(h))) setY(m[1], 'salario', c);
      else if ((m = /^personnel cost (\d{4})/.exec(h))) setY(m[1], 'pc', c);
      else if ((m = /^(\d{4}) pc para ratecard/.exec(h))) setY(m[1], 'custoOH', c);
      else if ((m = /^custo total (\d{4})/.exec(h))) setY(m[1], 'custo', c);
      else if ((m = /^ratecard.*\((\d{4})\)/.exec(h))) setY(m[1], 'rate', c);
    }
    const anos = Object.keys(yearCols).map(Number).filter(y => yearCols[y].rate || yearCols[y].custo).sort((a, b) => a - b);
    if (!anos.length) throw new Error('a aba "Custo Referência" não tem colunas de custo e ratecard por ano');
    const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
    const str = v => (v == null ? '' : String(v).trim());
    const perfis = []; const maxR = Math.max(...sh.rows.keys());
    for (let r = hr + 1; r <= maxR; r++) {
      const row = sh.rows.get(r); if (!row) continue;
      const g = c => { const x = c ? row.cells.get(c) : null; return x ? x.v : null; };
      const istudio = str(g(cols.istudio)), perfil = str(g(cols.perfil)), senioridade = str(g(cols.senioridade));
      if (!perfil || !senioridade) continue;
      const porAno = {};
      for (const y of anos) { const yc = yearCols[y]; const v = { salario: num(g(yc.salario)), pc: num(g(yc.pc)), custoOH: num(g(yc.custoOH)), custo: num(g(yc.custo)), rate: num(g(yc.rate)) }; if (v.rate != null || v.custo != null) porAno[y] = v; }
      perfis.push({ contrato: str(g(cols.contrato)), istudio, perfil, senioridade, completo: str(g(cols.completo)), anos: porAno, row: r, fonte: 'DRP' });
    }
    if (!perfis.length) throw new Error('a aba "Custo Referência" não tem perfis preenchidos');
    const uniq = a => [...new Set(a.filter(Boolean))].sort((x, y) => x.localeCompare(y));
    return { sheet: sheetName, perfis, anos, istudios: uniq(perfis.map(p => p.istudio)), contratos: uniq(perfis.map(p => p.contrato)) };
  }

  const api = { load, loadConfig, loadDRP, exportConfig, exportXlsx, calcAlloc, projectOf, refreshTrib, taxRate, shiftFormula, norm, serialToISO, isoToSerial, BL_INPUT, PROP_FIELDS, nid };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.SHCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
