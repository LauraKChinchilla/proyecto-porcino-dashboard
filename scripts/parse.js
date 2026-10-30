/* Lee los Excel originales del Proyecto Porcino y devuelve registros limpios.
   Funciona en Node (GitHub Actions) y en el navegador (pruebas).
   Es tolerante: busca hojas y columnas por su nombre, no por posición. */
(function (root) {
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const MESES_EN = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const pad = n => String(n).padStart(2, '0');
  const iso = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;

  function num(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).replace(/[^0-9.,-]/g, '');
    if (!s || s === '-' || s === '.') return null;
    const dots = (s.match(/\./g) || []).length, commas = (s.match(/,/g) || []).length;
    if (dots > 1) { const i = s.lastIndexOf('.'); s = s.slice(0, i).replace(/\./g, '') + '.' + s.slice(i + 1); }
    if (commas) s = s.replace(/,/g, dots ? '' : (commas === 1 && /,\d{1,2}$/.test(s) ? '.' : ''));
    const n = parseFloat(s);
    return isFinite(n) ? n : null;
  }
  // Fecha -> 'YYYY-MM-DD'. ref (iso) ayuda a adivinar el año de textos sin año.
  function toISO(v, ref) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) return iso(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === 'number' && v > 30000 && v < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5);
      return iso(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
    }
    const s = norm(v);
    let m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (m) { let y = +m[3]; if (y < 100) y += 2000; return iso(y, +m[1] > 12 ? +m[2] : +m[1], +m[1] > 12 ? +m[1] : +m[2]); }
    m = s.match(/(\d{1,2})\D{0,4}([a-z]{3,})\D*(\d{2,4})?/);
    if (m) {
      const k = m[2].slice(0, 3);
      let mi = MESES.indexOf(k); if (mi < 0) mi = MESES_EN.indexOf(k);
      if (mi < 0) mi = MESES.findIndex(x => x[0] === k[0] && x[1] === k[1]);
      if (mi >= 0) {
        let y = m[3] ? +m[3] : (ref ? +ref.slice(0, 4) : new Date().getFullYear());
        if (y < 100) y += 2000;
        let out = iso(y, mi + 1, +m[1]);
        if (!m[3] && ref && out < ref) out = iso(y + 1, mi + 1, +m[1]);
        return out;
      }
    }
    return null;
  }
  const addDays = (d, n) => { const t = new Date(d + 'T12:00:00Z'); t.setUTCDate(t.getUTCDate() + n); return t.toISOString().slice(0, 10); };
  const code = s => norm(s).replace(/cerda|cerd|cod/g, '').replace(/[^a-z0-9]/g, '').toUpperCase();

  function aoa(XLSX, ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true }); }
  function sheetBy(wb, ...keys) {
    const names = wb.SheetNames;
    for (const k of keys) { const n = names.find(x => norm(x).includes(k)); if (n) return wb.Sheets[n]; }
    return null;
  }
  function findHeader(rows, tests, maxRow = 30) {
    for (let r = 0; r < Math.min(rows.length, maxRow); r++) {
      const row = (rows[r] || []).map(norm);
      if (tests.every(t => row.some(c => t instanceof RegExp ? t.test(c) : c === t))) return r;
    }
    return -1;
  }
  const colOf = (row, test, from = 0) => row.findIndex((c, i) => i >= from && (test instanceof RegExp ? test.test(norm(c)) : norm(c) === test));

  // ---------- Categorías ----------
  function catLibro(txt, entrada) {
    const t = norm(txt);
    if (entrada) {
      if (/dinero en cuenta|saldo inicial/.test(t)) return 'Saldo inicial';
      if (/capital|aporte/.test(t)) return 'Aporte de socio';
      if (/venta/.test(t)) return 'Ventas';
      return 'Otros ingresos';
    }
    if (/comision|flete|viaje/.test(t)) return 'Flete / comisión';
    if (/concentrado|nutricon|nutri cerdo|bionova|gestacion|desarrollo|engorde|crecimiento|alcon/.test(t)) return 'Concentrado';
    if (/suagrovet|agromont|vacuna|antibiot|tylogan|yodo|cloro|podosept|bactrovet|guante|termometr|aguja|ahuja|catosal|ivermect|desparas|vitamin/.test(t)) return 'Sanidad e insumos';
    if (/teflon|disco|bebedero|cable|foco|broca|tornillo|ladrillo|pvc|pint|concreto|codo|cemento|bijao|sincel|cincel|brocha|tubo|tuvo|varilla|malla|hierro|electrodo|material/.test(t)) return 'Materiales y mantenimiento';
    if (/^compra$|^compra ·|sin descrip/.test(t) || t === 'compra') return 'Sin detalle';
    return 'Otros gastos';
  }
  function catInversion(txt) {
    const t = norm(txt);
    if (/restante|ajuste/.test(t) && /efectivo|dinero/.test(t)) return 'Ajuste';
    if (/arena/.test(t)) return 'Infraestructura y equipo';
    if (/cerda|padron|cerdit|verra|intergen|machito|lechon/.test(t)) return 'Animales';
    if (/concentr|bionova|nutri|bolsa|desarrollo|crecimiento|engorde|gestacion/.test(t)) return 'Concentrado';
    if (/antibiot|vacuna|catosal|ciperclor|chodine|dectoplus|bioquin|florpro|excenel|farro|aguja|ahuja/.test(t)) return 'Sanidad';
    if (/abogad/.test(t)) return 'Legal';
    if (/viaje|flete/.test(t)) return 'Fletes y viajes';
    if (/mano de obra|dias trabajo|dias de trabajo/.test(t)) return 'Mano de obra';
    return 'Infraestructura y equipo';
  }

  // ---------- Libro Diario ----------
  function parseLibro(XLSX, wb) {
    const ws = sheetBy(wb, 'libro diario', 'libro'); if (!ws) return { rows: [], warn: ['No se encontró la hoja Libro Diario'] };
    const rows = aoa(XLSX, ws);
    const h = findHeader(rows, ['entrada', 'salida']);
    if (h < 0) return { rows: [], warn: ['Libro Diario: no se encontraron las columnas Entrada y Salida'] };
    const H = rows[h];
    const cF = colOf(H, /^fecha/), cC = colOf(H, /^concepto/), cM = colOf(H, /^comentario/), cQ = colOf(H, /^cantidad/),
      cP = colOf(H, /^precio/), cE = colOf(H, 'entrada'), cS = colOf(H, 'salida');
    const out = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const ent = num(row[cE]) || 0, sal = num(row[cS]) || 0;
      if (!ent && !sal) continue;
      const parts = [row[cC], row[cM]].map(x => x == null ? '' : String(x).trim()).filter(Boolean);
      const concepto = parts.join(' · ') || '(sin concepto)';
      out.push({ fila: r + 1, fecha: cF >= 0 ? toISO(row[cF]) : null, concepto, cantidad: num(row[cQ]), precio: num(row[cP]),
        entrada: ent, salida: sal, cat: catLibro(concepto, ent > 0 && !sal) });
    }
    return { rows: out, warn: [] };
  }

  // ---------- Aporte Inicial ----------
  function parseAportes(XLSX, wb) {
    const ws = sheetBy(wb, 'aporte'); if (!ws) return null;
    const rows = aoa(XLSX, ws);
    const titles = [];
    rows.slice(0, 3).forEach(row => (row || []).forEach((c, i) => {
      const t = norm(c); if (/aporte/.test(t)) titles.push({ col: i, socio: /eric/.test(t) ? 'Erick' : /wili/.test(t) ? 'Wilian' : null });
    }));
    const h = findHeader(rows, [/^(activo|descri)/]);
    if (h < 0) return null;
    const blocks = [];
    rows[h].forEach((c, i) => {
      const t = norm(c); if (!/^(activo|descri)/.test(t)) return;
      const tt = titles.filter(x => x.col <= i + 1 && x.socio).sort((a, b) => b.col - a.col)[0];
      blocks.push({ desc: i, qty: i - 1, unit: i + 1, total: i + 2, fecha: i - 2, socio: tt ? tt.socio : null, compras: /^descri/.test(t) });
    });
    const items = [];
    for (const b of blocks) {
      if (!b.socio) continue;
      for (let r = h + 1; r < rows.length; r++) {
        const row = rows[r] || []; const d = row[b.desc];
        if (d == null || String(d).trim() === '' || /^descri/.test(norm(d))) continue;
        if (/falta|igualar/.test(norm(d))) continue;
        let tot = num(row[b.total]); const q = num(row[b.qty]), u = num(row[b.unit]);
        if (tot == null && q != null && u != null) tot = q * u;
        if (tot == null || tot === 0) continue;
        const efectivo = /efectivo|comision/.test(norm(d)) && !b.compras;
        items.push({ socio: b.socio, tipo: b.compras ? 'Compra con aporte' : (efectivo ? 'Efectivo' : 'Especie'),
          fecha: toISO(row[b.fecha]), descripcion: String(d).trim(), cantidad: q, unitario: u, total: tot, cat: catInversion(d) });
      }
    }
    return items;
  }

  // ---------- Celos y partos (Proyecto Porcino) ----------
  function parseCelos(XLSX, wb) {
    const ws = sheetBy(wb, 'celos'); if (!ws) return [];
    const rows = aoa(XLSX, ws);
    const h = findHeader(rows, ['animal', /^monta/]);
    if (h < 0) return [];
    const H = rows[h];
    const cA = colOf(H, 'animal'), cV = colOf(H, 'verraco'), cMonta = colOf(H, /^monta/), c114 = colOf(H, /114/),
      cDest = colOf(H, /^destete/), cNext = colOf(H, /^monta/, cMonta + 1), cC2 = colOf(H, /2.*celo/), cC3 = colOf(H, /3.*celo/);
    const out = []; let verraco = null;
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || []; const a = row[cA];
      if (cV >= 0 && row[cV] && !a) { continue; }
      if (cV >= 0 && row[cV]) verraco = String(row[cV]).trim();
      if (!a || !/cod|\d/.test(norm(a))) continue;
      const monta = toISO(row[cMonta]);
      const ref = monta || toISO(row[cC2]) || toISO(row[cC3]);
      let parto = c114 >= 0 ? toISO(row[c114], ref) : null;
      if (!parto && monta) parto = addDays(monta, 114);
      const prox = cNext >= 0 ? toISO(row[cNext], parto || ref) : null;
      out.push({ cerda: code(a), verraco: cV >= 0 && row[cV] ? String(row[cV]).trim() : null, monta, partoEsperado: parto,
        destete: cDest >= 0 ? toISO(row[cDest], parto || ref) : null, proximaMonta: prox, celo2: toISO(row[cC2], ref) });
    }
    return out;
  }

  // ---------- Vacunas ----------
  function parseVacunas(XLSX, wb) {
    const ws = sheetBy(wb, 'vacuna'); if (!ws) return [];
    const rows = aoa(XLSX, ws);
    const h = findHeader(rows, ['fecha', 'animal']); if (h < 0) return [];
    const H = rows[h]; const cF = colOf(H, 'fecha'), cA = colOf(H, 'animal');
    const out = []; let last = {};
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || []; if (!row[cA]) continue;
      const vac = row[cA + 1] ? String(row[cA + 1]).trim() : '';
      const dosis = row[cA + 2] && String(row[cA + 2]).trim().length > 0 && !/^"?\s*"?$/.test(String(row[cA + 2]).trim()) ? String(row[cA + 2]).trim() : last.dosis;
      const ref = toISO(row[cF]);
      const refuerzo = row[cA + 3] != null ? (toISO(row[cA + 3], ref) || String(row[cA + 3]).trim()) : null;
      const nota = row[cA + 4] != null && String(row[cA + 4]).trim().replace(/"/g, '').trim() ? String(row[cA + 4]).trim() : null;
      last = { dosis };
      out.push({ fecha: ref, cerda: code(row[cA]), vacuna: vac, dosis, refuerzo, nota });
    }
    return out;
  }

  // ---------- Sistema Contable ----------
  function parsePartos(XLSX, wb) {
    const ws = sheetBy(wb, 'parto'); if (!ws) return [];
    const rows = aoa(XLSX, ws);
    const h = findHeader(rows, [/fecha parto/]); if (h < 0) return [];
    const H = rows[h];
    const c = k => colOf(H, k);
    const cF = c(/fecha parto/), cN = c(/cerda/), cM = c(/fecha monta/), cT = c(/total nacidos/), cV = c(/nacidos vivos/),
      cMu = c(/nacidos muertos/), cB = c(/^bajas/), cP = c(/peso/), cO = c(/observ/), cD = c(/destete/), cMo = colOf(H, /^monta/);
    const out = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || []; if (!row[cN]) continue;
      const monta = toISO(row[cM]);
      const fecha = toISO(row[cF], monta);
      if (!fecha) continue;
      out.push({ cerda: code(row[cN]), fecha, monta, total: num(row[cT]), vivos: num(row[cV]), muertos: num(row[cMu]),
        bajas: num(row[cB]), peso: row[cP] != null ? String(row[cP]).trim() : null, obs: row[cO] ? String(row[cO]).trim() : null,
        destete: toISO(row[cD], fecha), proximaMonta: cMo >= 0 ? toISO(row[cMo], fecha) : null });
    }
    return out;
  }
  function parseCosto(XLSX, wb) {
    const ws = sheetBy(wb, 'costo'); if (!ws) return null;
    const rows = aoa(XLSX, ws);
    const h = findHeader(rows, [/dias uso/]);
    const items = []; let total = null;
    if (h >= 0) {
      const H = rows[h]; const cT = colOf(H, /^total/);
      for (let r = h + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        if (row.some(x => /costo total/.test(norm(x)))) { total = num(row[cT]); break; }
        const label = row.slice(0, cT).find(x => typeof x === 'string' && x.trim());
        const v = num(row[cT]);
        if (label && v) items.push({ concepto: String(label).trim(), total: v });
      }
    }
    let lechones = null, precio = null, ingresos = null;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      const cp = colOf(row, /precio de venta/);
      if (cp >= 0) {
        const nx = rows[r + 1] || [];
        precio = num(nx[cp]); lechones = num(nx[colOf(row, 'cantidad')]);
        const ci = colOf(row, /total ingresos/); ingresos = ci >= 0 ? num(nx[ci]) : null;
        const ct = colOf(row, 'total'); if (total == null && ct >= 0) total = num(nx[ct]);
        break;
      }
    }
    if (total == null) total = items.reduce((a, b) => a + b.total, 0);
    return { items, total, lechones, precio, ingresos };
  }
  function parseTabla(XLSX, wb, key, headTest) {
    const ws = sheetBy(wb, key); if (!ws) return [];
    const rows = aoa(XLSX, ws);
    const h = findHeader(rows, headTest); if (h < 0) return [];
    const H = rows[h].map(x => x == null ? '' : String(x).trim());
    const out = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || []; if (!row[0] || !String(row[0]).trim()) continue;
      const o = {}; H.forEach((k, i) => { if (k) o[k] = row[i] instanceof Date ? toISO(row[i]) : (row[i] == null ? null : (typeof row[i] === 'string' ? row[i].replace(/"/g, '').trim() : row[i])); });
      out.push(o);
    }
    return out;
  }

  function parseAll(XLSX, books) {
    const warn = [];
    const out = { libro: [], aportes: [], celos: [], vacunas: [], partos: [], costo: null, medicamentos: [], concentrado: [] };
    if (books.proyecto) {
      const l = parseLibro(XLSX, books.proyecto); out.libro = l.rows; warn.push(...l.warn);
      out.aportes = parseAportes(XLSX, books.proyecto) || [];
      out.celos = parseCelos(XLSX, books.proyecto);
      out.vacunas = parseVacunas(XLSX, books.proyecto);
    } else warn.push('No se encontró Proyecto Porcino.xlsx');
    if (books.contable) {
      out.partos = parsePartos(XLSX, books.contable);
      out.costo = parseCosto(XLSX, books.contable);
      out.medicamentos = parseTabla(XLSX, books.contable, 'medicamento', [/^descripcion/]).map(o => {
        const k = Object.keys(o);
        return { nombre: o[k[0]], presentacion: o[k[1]], cantidad: num(o[k[2]]), tipo: o[k[3]], vence: o[k[4]] };
      });
      out.concentrado = parseTabla(XLSX, books.contable, 'concentrado', [/^descripcion/]).map(o => {
        const k = Object.keys(o);
        return { tipo: o[k[0]], bolsas: num(o[k[1]]), compra: o[k[2]], vence: o[k[3]] };
      });
    } else warn.push('No se encontró Sistema_Contable_Granja_Porcina.xlsx');
    out.avisos = warn;
    return out;
  }

  const api = { parseAll, toISO, num, code, norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.PorcinoParse = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
