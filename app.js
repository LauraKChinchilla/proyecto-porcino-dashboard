/* Dashboard Granja Porcina — se alimenta de data.json, que GitHub Actions
   regenera desde los Excel de OneDrive. Todo cálculo con fechas se hace aquí,
   así "ayer" y "hoy" siempre son los de quien mira la página. */
(function () {
  const REPO = 'LauraKChinchilla/proyecto-porcino-dashboard';
  const WORKFLOW = 'actualizar-datos.yml';
  const GEST = 114, LACT = 28;
  const $ = id => document.getElementById(id);
  const L = n => 'L ' + Math.round(n || 0).toLocaleString('es-HN');
  const pct = n => (n * 100).toFixed(1).replace('.', ',') + '%';
  const dec = (n, d = 1) => Number(n || 0).toFixed(d).replace('.', ',');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const MESL = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const pad = n => String(n).padStart(2, '0');
  const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const addDays = (d, n) => { const t = new Date(d + 'T12:00:00Z'); t.setUTCDate(t.getUTCDate() + n); return t.toISOString().slice(0, 10); };
  const diffDays = (a, b) => Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 864e5);
  const fmtD = d => d ? `${+d.slice(8)} ${MES[+d.slice(5, 7) - 1]} ${d.slice(0, 4)}` : '—';
  const fmtDs = d => d ? `${+d.slice(8)} ${MES[+d.slice(5, 7) - 1]}` : '—';
  const monthKey = d => d.slice(0, 7);
  const monthLabel = k => `${MESL[+k.slice(5) - 1]} ${k.slice(2, 4)}`;
  const nextMonth = k => { let y = +k.slice(0, 4), m = +k.slice(5) + 1; if (m > 12) { m = 1; y++; } return `${y}-${pad(m)}`; };
  function ago(ts) {
    const s = (Date.now() - new Date(ts)) / 1000;
    if (s < 90) return 'hace un momento';
    if (s < 3600) return `hace ${Math.round(s / 60)} min`;
    if (s < 86400) return `hace ${Math.round(s / 3600)} h`;
    return `hace ${Math.round(s / 86400)} días`;
  }
  const fmtTS = ts => { const d = new Date(ts); return `${d.getDate()} ${MES[d.getMonth()]} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };

  let DATA = null, lastGen = null, S = {};

  /* ---------------- Cálculos ---------------- */
  function compute(D) {
    const T = todayISO(), Y = addDays(T, -1);
    const libro = D.libro;
    const fechas = libro.map(r => r.fecha).filter(Boolean).sort();
    const inicio = fechas[0] || Y;
    const dias = Math.max(1, diffDays(inicio, Y) + 1), meses = dias / 30.4;
    const ent = libro.reduce((a, r) => a + r.entrada, 0), sal = libro.reduce((a, r) => a + r.salida, 0);
    const ventas = libro.filter(r => r.cat === 'Ventas').reduce((a, r) => a + r.entrada, 0);

    // socios
    const soc = { Wilian: { especie: 0, efectivo: 0, igual: 0 }, Erick: { especie: 0, efectivo: 0, igual: 0 } };
    D.aportes.forEach(a => { if (!soc[a.socio]) return; if (a.tipo === 'Especie') soc[a.socio].especie += a.total; if (a.tipo === 'Efectivo') soc[a.socio].efectivo += a.total; });
    let igualSin = 0;
    libro.filter(r => r.cat === 'Aporte de socio').forEach(r => {
      const t = r.concepto.toLowerCase();
      if (/eric/.test(t)) soc.Erick.igual += r.entrada; else if (/wili/.test(t)) soc.Wilian.igual += r.entrada; else igualSin += r.entrada;
    });
    Object.values(soc).forEach(s => s.total = s.especie + s.efectivo + s.igual);

    // partos y eventos
    const partos = D.partos.slice().sort((a, b) => a.fecha < b.fecha ? -1 : 1);
    const reales = partos.filter(p => p.fecha <= T);
    const conDatos = reales.filter(p => p.vivos != null);
    const prom = conDatos.length ? conDatos.reduce((a, p) => a + p.vivos, 0) / conDatos.length : 10;
    const destDe = p => p.destete && p.destete > p.fecha ? p.destete : addDays(p.fecha, LACT);
    const destetadas = conDatos.filter(p => destDe(p) < T);
    const vivDest = destetadas.reduce((a, p) => a + p.vivos, 0), bajDest = destetadas.reduce((a, p) => a + (p.bajas || 0), 0);
    let mortC = null, vivC = 0, bajC = 0;

    const ev = [];
    const near = (c, d, tol = 25) => ev.some(e => e.cerda === c && Math.abs(diffDays(e.parto, d)) <= tol);
    reales.forEach(p => ev.push({ cerda: p.cerda, parto: p.fecha, tipo: 'real', vivos: p.vivos, bajas: p.bajas, destete: destDe(p) }));
    partos.filter(p => p.fecha > T).forEach(p => { if (!near(p.cerda, p.fecha)) ev.push({ cerda: p.cerda, parto: p.fecha, tipo: 'esperado', destete: addDays(p.fecha, LACT) }); });
    D.celos.forEach(c => { if (c.partoEsperado && !near(c.cerda, c.partoEsperado)) ev.push({ cerda: c.cerda, parto: c.partoEsperado, tipo: c.partoEsperado < T ? 'sin registrar' : (c.monta > T ? 'planificado' : 'esperado'), destete: addDays(c.partoEsperado, LACT), monta: c.monta }); });
    [...D.partos, ...D.celos].forEach(p => { if (p.proximaMonta) { const pe = addDays(p.proximaMonta, GEST); if (!near(p.cerda, pe)) ev.push({ cerda: p.cerda, parto: pe, tipo: 'planificado', destete: addDays(pe, LACT), monta: p.proximaMonta }); } });
    ev.sort((a, b) => a.parto < b.parto ? -1 : 1);

    // inventario de lechones por camada (hoja Control de lechones)
    const lotes = (D.lechones || []).map(l => {
      const reg = partos.find(p => l.cerdas.includes(p.cerda) && Math.abs(diffDays(p.fecha, l.fecha)) <= 10);
      const destete = reg && reg.destete && reg.destete > l.fecha ? reg.destete : addDays(l.fecha, LACT);
      const stock = Math.max(0, l.vivos - l.muertes - l.ventas);
      const edad = diffDays(l.fecha, T);
      return { ...l, nombre: l.cerdas.join(' y ') || '—', destete, stock, edad, listo: destete <= T, reg };
    });
    const listos = lotes.filter(l => l.listo).reduce((a, l) => a + l.stock, 0);
    const lactando = lotes.filter(l => !l.listo).reduce((a, l) => a + l.stock, 0);
    // mortalidad: preferir Control de lechones (se lleva al día); si no hay, usar el registro de partos
    const lDest = lotes.filter(l => l.listo && l.vivos);
    if (lDest.length) {
      const v = lDest.reduce((a, l) => a + l.vivos, 0), m = lDest.reduce((a, l) => a + l.muertes, 0);
      if (v) { mortC = m / v; vivC = v; bajC = m; }
    }

    // inventario vacunas próximas
    const refuerzos = D.vacunas.filter(v => /^\d{4}-/.test(v.refuerzo || '') && v.refuerzo >= addDays(T, -3)).sort((a, b) => a.refuerzo < b.refuerzo ? -1 : 1);

    return { T, Y, inicio, dias, meses, ent, sal, saldo: ent - sal, ventas, gastoMes: sal / meses, ventaMes: ventas / meses, soc, igualSin,
      partos, reales, conDatos, prom, ev, refuerzos, lotes, listos, lactando,
      mort: mortC != null ? mortC : (vivDest ? bajDest / vivDest : 0), vivDest: mortC != null ? vivC : vivDest, bajDest: mortC != null ? bajC : bajDest,
      mortFuente: mortC != null ? 'Control de lechones' : 'Registro de partos' };
  }

  /* ---------------- Render ---------------- */
  function bars(el, rows, color, total) {
    const max = Math.max(1, ...rows.map(r => Math.abs(r[1])));
    el.innerHTML = rows.map(r => `<div class="hbar" title="${esc(r[0])}: ${L(r[1])}${total ? ' (' + pct(r[1] / total) + ')' : ''}"><span class="n">${esc(r[0])}</span><span class="t"><span class="f" style="display:block;width:${(Math.abs(r[1]) / max * 100).toFixed(1)}%;background:${color}"></span></span><span class="v">${L(r[1])}${total ? ` <span class="muted">${Math.round(r[1] / total * 100)}%</span>` : ''}</span></div>`).join('') || '<p class="muted">Sin datos</p>';
  }
  const groupBy = (arr, key, val) => Object.entries(arr.reduce((m, r) => { const k = key(r), v = val(r); if (v) m[k] = (m[k] || 0) + v; return m; }, {})).sort((a, b) => b[1] - a[1]);

  function render() {
    const D = DATA, C = S = compute(D);
    $('periodo').innerHTML = `Periodo del libro diario: <b>${fmtD(C.inicio)}</b> → <b>${fmtD(C.Y)}</b> · ${C.dias} días`;

    // KPIs
    const cap = C.soc.Wilian.total + C.soc.Erick.total, dif = C.soc.Wilian.total - C.soc.Erick.total;
    const def = C.gastoMes - C.ventaMes;
    const precio = (D.costo && D.costo.precio) || 2300;
    const costoU = D.costo && D.costo.lechones ? D.costo.total / D.costo.lechones : null;
    const kpis = [
      ['Capital total de socios', L(cap), `Wilian ${L(C.soc.Wilian.total)} · Erick ${L(C.soc.Erick.total)}`, Math.abs(dif) < 1 ? ['good', '✓ Igualados'] : ['warn', '! Diferencia ' + L(Math.abs(dif))]],
      ['Saldo en caja', L(C.saldo), `Entradas ${L(C.ent)} − salidas ${L(C.sal)}`, C.saldo < 0 ? ['bad', '▼ Negativo'] : null],
      ['Gasto promedio mensual', L(C.gastoMes), `${L(C.sal / C.dias)} por día`, null],
      ['Ventas promedio mensuales', L(C.ventaMes), `${L(C.ventas)} vendidos en el periodo`, null],
      ['Resultado operativo', L(C.ventas - C.sal), `Ventas − gastos en ${C.dias} días`, C.ventas - C.sal < 0 ? ['bad', '▼ Déficit'] : ['good', '▲ Superávit']],
      ['Cobertura de caja', def > 0 ? dec(C.saldo / def) + ' meses' : 'Sin déficit', 'Al ritmo actual de gasto y venta', def > 0 && C.saldo / def < 2 ? ['warn', '! Ajustado'] : null],
      ['Listos para vender', `${C.listos} lechones`, `≈ ${L(C.listos * precio)} a ${L(precio)} · ${C.lactando} más en lactancia`, C.listos ? ['good', '● Destetados'] : null],
      ['Mortalidad de lechones', pct(C.mort), `${C.bajDest} muertes de ${C.vivDest} vivos (camadas destetadas)`, C.mort > 0.2 ? ['bad', '▲ Alta'] : ['good', '✓ Aceptable']],
      ['Margen por lechón', costoU != null ? L(precio - costoU) : '—', costoU != null ? `Costo ${L(costoU)} · precio ${L(precio)}` : 'Sin hoja de costo', costoU != null && precio > costoU ? ['good', '✓ Positivo'] : null],
    ];
    $('kpiGrid').innerHTML = kpis.map(k => `<div class="card kpi"><div class="lbl">${k[0]}</div><div class="val">${k[1]}</div><div class="note">${k[2]}</div>${k[3] ? `<div style="margin-top:8px"><span class="badge ${k[3][0]}">${k[3][1]}</span></div>` : ''}</div>`).join('');

    // Flujo
    $('entTot').textContent = L(C.ent); $('salTot').textContent = L(C.sal);
    bars($('entBars'), groupBy(D.libro, r => r.cat, r => r.entrada), 'var(--s3)', C.ent);
    bars($('salBars'), groupBy(D.libro, r => r.cat, r => r.salida), 'var(--s2)', C.sal);
    $('libro').innerHTML = '<thead><tr><th>Fila</th><th>Fecha</th><th>Concepto</th><th>Categoría</th><th class="num">Entrada</th><th class="num">Salida</th></tr></thead><tbody>' +
      D.libro.map(r => `<tr><td class="muted">${r.fila}</td><td class="muted" style="white-space:nowrap">${r.fecha ? fmtDs(r.fecha) : '<span class="badge warn">sin fecha</span>'}</td><td>${esc(r.concepto)}</td><td class="muted">${esc(r.cat)}</td><td class="num">${r.entrada ? L(r.entrada) : ''}</td><td class="num">${r.salida ? L(r.salida) : ''}</td></tr>`).join('') +
      `<tr class="tot"><td></td><td></td><td>Total</td><td></td><td class="num">${L(C.ent)}</td><td class="num">${L(C.sal)}</td></tr><tr class="tot"><td></td><td></td><td>Saldo</td><td></td><td class="num" colspan="2">${L(C.saldo)}</td></tr></tbody>`;

    // Socios
    const socs = [['Wilian', C.soc.Wilian], ['Erick', C.soc.Erick]];
    $('socioSplit').innerHTML = socs.map((s, i) => `<div title="${s[0]}: ${L(s[1].total)}" style="width:${cap ? s[1].total / cap * 100 : 50}%;background:var(--s${i + 1})"></div>`).join('');
    $('socioLegend').innerHTML = socs.map((s, i) => `<span><i style="background:var(--s${i + 1})"></i>${s[0]} · ${L(s[1].total)} (${cap ? Math.round(s[1].total / cap * 100) : 0}%)</span>`).join('');
    $('socioTbl').innerHTML = '<thead><tr><th>Socio</th><th class="num">Especie</th><th class="num">Efectivo</th><th class="num">Igualación</th><th class="num">Total</th></tr></thead><tbody>' +
      socs.map(s => `<tr><td>${s[0]}</td><td class="num">${L(s[1].especie)}</td><td class="num">${L(s[1].efectivo)}</td><td class="num">${L(s[1].igual)}</td><td class="num"><b>${L(s[1].total)}</b></td></tr>`).join('') +
      `<tr class="tot"><td>Total</td><td class="num">${L(C.soc.Wilian.especie + C.soc.Erick.especie)}</td><td class="num">${L(C.soc.Wilian.efectivo + C.soc.Erick.efectivo)}</td><td class="num">${L(C.soc.Wilian.igual + C.soc.Erick.igual)}</td><td class="num">${L(cap)}</td></tr></tbody>`;
    const inv = D.aportes.filter(a => a.tipo !== 'Efectivo' && a.cat !== 'Ajuste');
    const cats = groupBy(inv, a => a.cat, a => a.total).map(c => c[0]);
    const invRows = cats.map(c => [c, inv.filter(a => a.cat === c && a.socio === 'Wilian').reduce((x, a) => x + a.total, 0), inv.filter(a => a.cat === c && a.socio === 'Erick').reduce((x, a) => x + a.total, 0)]);
    const invMax = Math.max(1, ...invRows.map(r => r[1] + r[2]));
    const ajuste = D.aportes.filter(a => a.cat === 'Ajuste').reduce((x, a) => x + a.total, 0);
    $('invBars').innerHTML = invRows.map(r => `<div class="hbar" title="${esc(r[0])}: Wilian ${L(r[1])} · Erick ${L(r[2])}"><span class="n">${esc(r[0])}</span><span class="t" style="display:flex;gap:2px;background:transparent">${r[1] ? `<span style="width:${r[1] / invMax * 100}%;background:var(--s1);border-radius:4px"></span>` : ''}${r[2] ? `<span style="width:${r[2] / invMax * 100}%;background:var(--s2);border-radius:4px"></span>` : ''}</span><span class="v">${L(r[1] + r[2])}</span></div>`).join('') +
      (ajuste ? `<p class="cap" style="margin:10px 0 0">Incluye un ajuste de ${L(ajuste)} registrado por Wilian (“efectivo restante”), ya descontado de su total.</p>` : '');

    // Producción
    const pMax = Math.max(1, ...C.reales.map(p => p.total || p.vivos || 0));
    $('partoBars').innerHTML = C.reales.map(p => {
      if (p.vivos == null) return `<div class="hbar"><span class="n">${esc(p.cerda)} · ${fmtDs(p.fecha)}</span><span class="t"></span><span class="v"><span class="badge warn">sin datos</span></span></div>`;
      const b = p.bajas || 0, sob = Math.max(0, p.vivos - b), m = p.muertos || 0;
      return `<div class="hbar" title="Cerda ${esc(p.cerda)} (${fmtD(p.fecha)}): ${p.total ?? p.vivos} nacidos, ${p.vivos} vivos, ${b} bajas"><span class="n">${esc(p.cerda)} · ${fmtDs(p.fecha)}</span><span class="t" style="display:flex;gap:2px;background:transparent">${sob ? `<span style="width:${sob / pMax * 100}%;background:var(--s3);border-radius:4px"></span>` : ''}${b ? `<span style="width:${b / pMax * 100}%;background:var(--s2);border-radius:4px"></span>` : ''}${m ? `<span style="width:${m / pMax * 100}%;background:var(--muted);border-radius:4px"></span>` : ''}</span><span class="v">${sob} de ${p.total ?? p.vivos}</span></div>`;
    }).join('') + `<p class="cap" style="margin:10px 0 0">${C.conDatos.reduce((a, p) => a + (p.total || p.vivos), 0)} nacidos · ${C.conDatos.reduce((a, p) => a + p.vivos, 0)} vivos · ${C.conDatos.reduce((a, p) => a + (p.bajas || 0), 0)} bajas · promedio ${dec(C.prom)} vivos por camada.</p>`;
    if (D.costo) {
      const k = D.costo;
      $('costoTbl').innerHTML = '<tbody>' + k.items.map(i => `<tr><td>${esc(i.concepto.charAt(0).toUpperCase() + i.concepto.slice(1))}</td><td class="num">${L(i.total)}</td></tr>`).join('') +
        `<tr class="tot"><td>Costo por camada</td><td class="num">${L(k.total)}</td></tr><tr><td>Costo por lechón (÷${k.lechones})</td><td class="num">${L(costoU)}</td></tr><tr><td>Precio de venta</td><td class="num">${L(precio)}</td></tr></tbody>`;
      const vend = C.prom * (1 - C.mort);
      $('margenCallout').innerHTML = `Cada lechón vendido deja <b>${L(precio - costoU)}</b> según el modelo. Con la mortalidad real (<b>${pct(C.mort)}</b>), una camada de ${dec(C.prom)} vivos termina en ~${dec(vend)} lechones vendibles, y el costo real sube a <b>${L(vend ? k.total / vend : 0)}</b> por lechón${vend && k.total / vend > precio ? ' — <b>por encima del precio de venta</b>' : ''}.`;
    }

    // Lechones en granja
    const lMax = Math.max(1, ...C.lotes.map(l => l.vivos));
    $('loteBars').innerHTML = C.lotes.length ? C.lotes.map(l => {
      const seg = (n, col, t) => n ? `<span title="${t}: ${n}" style="width:${n / lMax * 100}%;background:${col};border-radius:4px"></span>` : '';
      return `<div class="hbar" title="Cerda ${esc(l.nombre)} · parto ${fmtD(l.fecha)} · ${l.vivos} vivos, ${l.muertes} muertes, ${l.ventas} vendidos, ${l.stock} en granja"><span class="n">${esc(l.nombre)} · ${fmtDs(l.fecha)}</span><span class="t" style="display:flex;gap:2px;background:transparent">${seg(l.ventas, 'var(--s1)', 'Vendidos')}${seg(l.muertes, 'var(--s2)', 'Muertes')}${seg(l.stock, l.listo ? 'var(--s3)' : 'var(--s4)', l.listo ? 'Listos para vender' : 'En lactancia')}</span><span class="v">${l.stock ? `<b>${l.stock}</b> ${l.listo ? 'listos' : 'lactando'}` : '<span class="muted">0 en granja</span>'}</span></div>`;
    }).join('') : '<p class="muted">La hoja “Control de lechones” no tiene registros.</p>';
    $('loteTbl').innerHTML = '<thead><tr><th>Camada</th><th>Parto</th><th class="num">Edad</th><th class="num">En granja</th><th>Estado</th></tr></thead><tbody>' +
      C.lotes.filter(l => l.stock).map(l => `<tr><td>${esc(l.nombre)}</td><td>${fmtDs(l.fecha)}</td><td class="num">${l.edad} días</td><td class="num"><b>${l.stock}</b></td><td>${l.listo ? '<span class="badge good">Listo para vender</span>' : `<span class="badge warn">Destete ${fmtDs(l.destete)}</span>`}</td></tr>`).join('') +
      `<tr class="tot"><td>Total</td><td></td><td></td><td class="num">${C.listos + C.lactando}</td><td>${C.listos} listos · ${C.lactando} lactando</td></tr></tbody>`;
    $('loteCallout').innerHTML = C.listos ? `Hay <b>${C.listos} lechones destetados sin vender</b>: a ${L(precio)} son <b>${L(C.listos * precio)}</b>. La proyección los cuenta como venta en el próximo mes.` : 'No hay lechones destetados pendientes de venta según la hoja Control de lechones.';

    // Calendario
    const desde = monthKey(addDays(C.T, -100)), hasta = monthKey(addDays(C.T, 185));
    const meses = []; for (let k = desde; k <= hasta; k = nextMonth(k)) meses.push(k);
    const cls = { real: 'done', esperado: '', planificado: 'plan', 'sin registrar': 'late' };
    $('timeline').innerHTML = meses.map(k => {
      const es = C.ev.filter(e => monthKey(e.parto) === k);
      return `<div class="tl-row"><div class="tl-m">${monthLabel(k)}<div class="muted" style="font-weight:400;font-size:12px">${es.length} parto${es.length === 1 ? '' : 's'}</div></div><div class="chips">${es.length ? es.map(e => `<span class="chip ${cls[e.tipo]}" title="${e.tipo}${e.monta ? ' · monta ' + fmtD(e.monta) : ''}"><b>${esc(e.cerda)}</b> · ${+e.parto.slice(8)}</span>`).join('') : `<span class="muted" style="font-size:13px">Sin partos${k > monthKey(C.T) ? ': mes sin destetes al siguiente' : ''}</span>`}</div></div>`;
    }).join('');

    // Inventario
    $('medTbl').innerHTML = '<thead><tr><th>Producto</th><th>Tipo</th><th class="num">Frascos</th><th>Vence</th><th>Estado</th></tr></thead><tbody>' + D.medicamentos.map(m => {
      const v = /^\d{4}-\d{2}/.test(m.vence || '') ? m.vence : null;
      const st = !v ? ['', 'Sin fecha'] : v < C.T ? ['bad', 'Vencido'] : diffDays(C.T, v) <= 90 ? ['warn', 'Por vencer'] : ['good', 'Vigente'];
      const agot = m.cantidad === 0;
      return `<tr><td>${esc(m.nombre)}</td><td class="muted">${esc(m.tipo || '')}</td><td class="num">${m.cantidad ?? '—'}</td><td>${v ? MESL[+v.slice(5, 7) - 1] + ' ' + v.slice(0, 4) : '—'}</td><td>${agot ? '<span class="badge warn">Agotado</span> ' : ''}<span class="badge ${st[0]}">${st[1]}</span></td></tr>`;
    }).join('') + '</tbody>';
    $('concTbl').innerHTML = '<thead><tr><th>Concentrado</th><th class="num">Bolsas</th><th>Vence</th></tr></thead><tbody>' + D.concentrado.map(c => {
      const v = /^\d{4}-/.test(c.vence || '') ? c.vence : null;
      return `<tr><td>${esc(c.tipo)}</td><td class="num">${c.bolsas ?? '—'}${c.bolsas != null && c.bolsas <= 1 ? ' <span class="badge warn">bajo</span>' : ''}</td><td>${v ? fmtD(v) + (v < C.T ? ' <span class="badge bad">vencido</span>' : '') : '—'}</td></tr>`;
    }).join('') + '</tbody>';
    $('vacTbl').innerHTML = C.refuerzos.length ? '<thead><tr><th>Cerda</th><th>Vacuna</th><th>Refuerzo</th></tr></thead><tbody>' + C.refuerzos.map(v => `<tr><td>${esc(v.cerda)}</td><td>${esc(v.vacuna)}</td><td>${fmtD(v.refuerzo)}${v.refuerzo < C.T ? ' <span class="badge warn">¿aplicado?</span>' : ''}</td></tr>`).join('') + '</tbody>' : '<tbody><tr><td class="muted">No hay refuerzos con fecha próxima en el registro de vacunas.</td></tr></tbody>';

    // Notas automáticas
    const notas = [];
    const sinFecha = D.libro.filter(r => !r.fecha).length;
    if (sinFecha) notas.push(`${sinFecha} de ${D.libro.length} movimientos del Libro Diario no tienen fecha; por eso el gasto se calcula como promedio del periodo y no mes a mes.`);
    const sinDet = D.libro.filter(r => r.cat === 'Sin detalle');
    if (sinDet.length) notas.push(`${sinDet.length} gasto(s) sin descripción (${sinDet.map(r => L(r.salida)).join(', ')}).`);
    if (C.igualSin) notas.push(`Hay ${L(C.igualSin)} de ingreso de capital sin socio identificado.`);
    C.reales.filter(p => p.vivos == null).forEach(p => notas.push(`La cerda ${p.cerda} parió el ${fmtD(p.fecha)} pero no tiene nacidos ni vivos registrados.`));
    C.ev.filter(e => e.tipo === 'sin registrar').forEach(e => notas.push(`La cerda ${e.cerda} debía parir el ${fmtD(e.parto)} y no aparece en el registro de partos.`));
    C.partos.filter(p => p.fecha <= C.T && p.vivos != null && p.bajas == null && diffDays(p.fecha, C.T) > 7).forEach(p => notas.push(`Cerda ${p.cerda}: no se han anotado bajas desde el parto (${fmtD(p.fecha)}). Si no hubo, escribir 0.`));
    C.lotes.forEach(l => { if (l.reg && l.reg.vivos != null && l.cerdas.length === 1 && l.reg.vivos !== l.vivos) notas.push(`Cerda ${l.nombre}: ${l.vivos} vivos en “Control de lechones” pero ${l.reg.vivos} en “Registro de parto”. Se usa Control de lechones para el inventario.`); });
    if (Math.abs(dif) >= 1) notas.push(`Los aportes de los socios no están igualados: diferencia de ${L(Math.abs(dif))}.`);
    (D.avisos || []).forEach(a => notas.push(a));
    $('notasAuto').innerHTML = notas.length ? notas.map(n => `<li>${esc(n)}</li>`).join('') : '<li>Sin observaciones.</li>';

    // Proyección: sliders con defaults reales la primera vez
    const m = $('mort'), g = $('gasto'), p = $('precio');
    if (!m.dataset.touched) m.value = Math.round(C.mort * 100);
    if (!g.dataset.touched) g.value = Math.round(C.gastoMes / 500) * 500;
    if (!p.dataset.touched) p.value = precio;
    proj();
  }

  function proj() {
    const C = S; if (!C.T) return;
    const m = +$('mort').value / 100, p = +$('precio').value, g = +$('gasto').value;
    $('mV').textContent = Math.round(m * 100) + '%'; $('pV').textContent = L(p); $('gV').textContent = L(g);
    let k = nextMonth(monthKey(C.T)); const meses = []; for (let i = 0; i < 6; i++) { meses.push(k); k = nextMonth(k); }
    let caja = C.saldo;
    const rows = meses.map((mk, i) => {
      let lech = i === 0 ? C.listos : 0, cam = 0;
      C.ev.forEach(e => {
        if (monthKey(e.destete) !== mk || e.destete < C.T) return;
        cam++;
        const lote = e.tipo === 'real' && C.lotes.find(l => l.cerdas.includes(e.cerda) && Math.abs(diffDays(l.fecha, e.parto)) <= 10);
        if (lote) { const rest = Math.min(1, Math.max(0, diffDays(C.T, lote.destete)) / LACT); lech += (lote.cerdas.length > 1 ? lote.stock / lote.cerdas.length : lote.stock) * (1 - m * rest); }
        else if (e.tipo === 'real' && e.vivos != null) lech += Math.min(e.vivos - (e.bajas || 0), e.vivos * (1 - m));
        else lech += C.prom * (1 - m);
      });
      const ing = Math.max(0, lech) * p; caja += ing - g;
      return { mes: monthLabel(mk), cam, lech: Math.max(0, lech), ing, g, caja };
    });
    draw(rows);
    $('projTbl').innerHTML = '<thead><tr><th>Mes</th><th class="num">Camadas al destete</th><th class="num">Lechones a la venta</th><th class="num">Ingresos</th><th class="num">Gasto</th><th class="num">Neto</th><th class="num">Caja acumulada</th></tr></thead><tbody>' +
      rows.map(r => `<tr><td>${r.mes}</td><td class="num">${r.cam}</td><td class="num">${Math.round(r.lech)}</td><td class="num">${L(r.ing)}</td><td class="num">${L(r.g)}</td><td class="num" style="color:${r.ing - r.g < 0 ? 'var(--bad)' : 'var(--good)'}">${L(r.ing - r.g)}</td><td class="num"><b>${L(r.caja)}</b></td></tr>`).join('') + '</tbody>';
    const neg = rows.filter(r => r.caja < 0), fin = rows[rows.length - 1];
    const tot = rows.reduce((a, r) => a + r.ing, 0), totG = rows.reduce((a, r) => a + r.g, 0);
    const low = rows.reduce((a, r) => r.caja < a.caja ? r : a, rows[0]);
    $('projCallout').innerHTML = (neg.length ? `⚠ Con estos supuestos la caja <b>queda en negativo en ${neg.map(r => r.mes).join(', ')}</b>; lo más bajo es ${low.mes} (${L(low.caja)}). Hay que cubrir ese hueco con aportes, venta de cerdos de engorde o recorte de gasto. ` : '✓ Con estos supuestos la caja <b>no baja de cero</b> en los seis meses. ') +
      (C.listos ? `${monthLabel(nextMonth(monthKey(C.T)))} incluye los ${C.listos} lechones ya destetados. ` : '') + `En seis meses entrarían <b>${L(tot)}</b> por lechones contra <b>${L(totG)}</b> de gasto; caja proyectada al final: <b>${L(fin.caja)}</b>. Solo cuenta montas y partos que ya están en los Excel.`;
  }

  function draw(rows) {
    const W = 900, H = 320, pl = 64, pr = 16, pt = 22, pb = 34, el = $('projChart');
    const vals = rows.flatMap(r => [r.ing, r.g, r.caja]); const max = Math.max(...vals, 0) * 1.1 || 1, min = Math.min(...vals, 0) * 1.1;
    const y = v => pt + (max - v) / (max - min) * (H - pt - pb); const bw = (W - pl - pr) / rows.length;
    const step = Math.pow(10, Math.floor(Math.log10((max - min) / 4))); const nice = Math.ceil((max - min) / 4 / step) * step;
    let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Proyección de caja">`;
    for (let t = Math.ceil(min / nice) * nice; t <= max; t += nice) s += `<line x1="${pl}" x2="${W - pr}" y1="${y(t)}" y2="${y(t)}" stroke="${t === 0 ? 'var(--muted)' : 'var(--line)'}"/><text x="${pl - 8}" y="${y(t) + 4}" text-anchor="end" font-size="11" fill="var(--muted)">${t === 0 ? '0' : Math.round(t / 1000).toLocaleString('es-HN') + 'k'}</text>`;
    rows.forEach((r, i) => {
      const x = pl + i * bw, w = Math.min(26, bw * .24);
      s += `<rect x="${x + bw / 2 - w - 1}" y="${y(r.ing)}" width="${w}" height="${Math.max(0, y(0) - y(r.ing))}" rx="3" fill="var(--s3)"/>`;
      s += `<rect x="${x + bw / 2 + 1}" y="${y(r.g)}" width="${w}" height="${Math.max(0, y(0) - y(r.g))}" rx="3" fill="var(--s2)"/>`;
      s += `<text x="${x + bw / 2}" y="${H - 10}" text-anchor="middle" font-size="12" fill="var(--text-2)">${r.mes}</text>`;
      s += `<rect class="hit" data-i="${i}" x="${x}" y="${pt}" width="${bw}" height="${H - pt - pb}" fill="transparent"/>`;
    });
    const pts = rows.map((r, i) => [pl + i * bw + bw / 2, y(r.caja)]);
    s += `<polyline points="${pts.map(p => p.join(',')).join(' ')}" fill="none" stroke="var(--s1)" stroke-width="2.5" stroke-linejoin="round"/>`;
    pts.forEach(p => { s += `<circle cx="${p[0]}" cy="${p[1]}" r="5" fill="var(--s1)" stroke="var(--surface)" stroke-width="2" pointer-events="none"/>`; });
    const last = pts[pts.length - 1]; s += `<text x="${last[0]}" y="${last[1] - 12}" text-anchor="middle" font-size="12" font-weight="600" fill="var(--text)">${L(rows[rows.length - 1].caja)}</text></svg>`;
    const tip = $('tip'); el.querySelectorAll('svg').forEach(n => n.remove()); el.insertAdjacentHTML('beforeend', s);
    el.querySelectorAll('.hit').forEach(h => {
      h.addEventListener('mousemove', e => {
        const r = rows[+h.dataset.i], b = el.getBoundingClientRect();
        tip.innerHTML = `<b>${r.mes}</b><br>${r.cam} camada(s) · ${Math.round(r.lech)} lechones<br>Ingresos ${L(r.ing)} · Gasto ${L(r.g)}<br>Caja ${L(r.caja)}`;
        tip.style.opacity = 1; let x = e.clientX - b.left + 12; if (x > b.width - 200) x -= 220; tip.style.left = x + 'px'; tip.style.top = (e.clientY - b.top - 10) + 'px';
      });
      h.addEventListener('mouseleave', () => tip.style.opacity = 0);
    });
  }

  /* ---------------- Frescura de datos ---------------- */
  let lastRun = null, live = false, liveSince = null;
  function renderFresh() {
    if (!DATA) return;
    const mods = (DATA.archivos || []).map(a => a.modificado).filter(Boolean).sort();
    const lastMod = mods[mods.length - 1];
    if (live) $('fRead').innerHTML = `<b>hace menos de 2 min</b> <span class="muted">· lectura continua activa desde ${fmtTS(liveSince)}</span>`;
    else $('fRead').innerHTML = lastRun ? `<b>${ago(lastRun)}</b> <span class="muted">(${fmtTS(lastRun)})</span>` : '<span class="muted">consultando…</span>';
    $('fMod').innerHTML = lastMod ? `<b>${ago(lastMod)}</b> <span class="muted">(${fmtTS(lastMod)})</span>` : '—';
    $('fFiles').innerHTML = (DATA.archivos || []).map(a => `${esc(a.nombre)} · ${fmtTS(a.modificado)}`).join('<br>');
    const stale = !live && lastRun && (Date.now() - new Date(lastRun)) > 30 * 60e3;
    $('liveDot').className = 'dot ' + (stale ? 'stale' : 'live');
    $('liveTxt').textContent = stale ? 'Lectura atrasada' : 'En vivo';
  }
  async function checkRun() {
    try {
      const r = await fetch(`https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}/runs?per_page=10`, { cache: 'no-store' });
      if (r.ok) {
        const runs = (await r.json()).workflow_runs || [];
        const act = runs.find(x => x.status === 'in_progress');
        live = !!act; liveSince = act ? (act.run_started_at || act.created_at) : null;
        const ok = runs.find(x => x.status === 'completed' && x.conclusion === 'success');
        if (ok) lastRun = ok.updated_at;
      }
    } catch (e) { /* sin acceso a la API: se muestra la fecha de los datos */ }
    if (!lastRun && DATA) lastRun = DATA.generado;
    renderFresh();
  }
  async function load() {
    try {
      const r = await fetch('data.json?t=' + Date.now(), { cache: 'no-store' });
      if (!r.ok) throw new Error(r.status);
      const d = await r.json();
      if (d.generado !== lastGen) { DATA = d; lastGen = d.generado; render(); $('err').hidden = true; }
      renderFresh();
    } catch (e) { $('err').hidden = false; }
  }

  ['mort', 'precio', 'gasto'].forEach(id => $(id).addEventListener('input', e => { e.target.dataset.touched = 1; proj(); }));
  $('themeBtn').onclick = () => { const r = document.documentElement; const dark = r.dataset.theme ? r.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches; r.dataset.theme = dark ? 'light' : 'dark'; };
  load(); checkRun();
  setInterval(load, 60e3);          // revisa si hay datos nuevos cada minuto
  setInterval(checkRun, 5 * 60e3);  // hora de la última lectura cada 5 min
  setInterval(renderFresh, 30e3);   // refresca el "hace X min"
})();
