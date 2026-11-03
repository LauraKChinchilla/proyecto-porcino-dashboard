/* Pantalla administrativa: guarda capturas como archivos en pendientes/ del repositorio.
   GitHub Actions las pasa al Excel de OneDrive. El token de GitHub vive cifrado
   (AES-GCM, clave derivada de la contraseña con PBKDF2) en admin/config.json. */
(function () {
  const OWNER = 'LauraKChinchilla', REPO = 'proyecto-porcino-dashboard', API = `https://api.github.com/repos/${OWNER}/${REPO}`;
  const $ = id => document.getElementById(id);
  const qs = (s, r = document) => r.querySelector(s), qsa = (s, r = document) => [...r.querySelectorAll(s)];
  const pad = n => String(n).padStart(2, '0');
  const hoy = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const addDays = (d, n) => { const t = new Date(d + 'T12:00:00'); t.setDate(t.getDate() + n); return `${t.getFullYear()}-${pad(t.getMonth() + 1)}-${pad(t.getDate())}`; };
  const MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const fmtD = d => d ? `${+d.slice(8, 10)} ${MES[+d.slice(5, 7) - 1]}` : '';
  const L = n => 'L ' + Math.round(n || 0).toLocaleString('es-HN');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const store = { get(k) { try { return localStorage.getItem(k) || sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v, persist) { try { (persist ? localStorage : sessionStorage).setItem(k, v); } catch (e) { } },
    del(k) { try { localStorage.removeItem(k); sessionStorage.removeItem(k); } catch (e) { } } };

  let TOKEN = null, DATA = null, CONFIG = null;

  function toast(msg, err) { const t = $('toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : ''); clearTimeout(t._h); t._h = setTimeout(() => t.className = 'toast', err ? 6000 : 3500); }

  /* ---------- cifrado ---------- */
  const enc = new TextEncoder(), dec = new TextDecoder();
  const b64 = buf => btoa(String.fromCharCode(...new Uint8Array(buf)));
  const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  async function clave(pass, salt, iter) {
    const base = await crypto.subtle.importKey('raw', enc.encode(pass), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }
  async function cifrar(texto, pass) {
    const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12)), iter = 310000;
    const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await clave(pass, salt, iter), enc.encode(texto));
    return { v: 1, iter, salt: b64(salt), iv: b64(iv), ct: b64(ct), creado: new Date().toISOString() };
  }
  async function descifrar(cfg, pass) {
    const k = await clave(pass, unb64(cfg.salt), cfg.iter);
    return dec.decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(cfg.iv) }, k, unb64(cfg.ct)));
  }

  /* ---------- GitHub ---------- */
  async function gh(path, opts = {}, tok = TOKEN) {
    const r = await fetch(API + path, { ...opts, cache: 'no-store', headers: { Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28', ...(opts.headers || {}) } });
    if (!r.ok && r.status !== 404) throw new Error(`GitHub respondió ${r.status}`);
    return r;
  }
  const utf8b64 = s => btoa(unescape(encodeURIComponent(s)));

  /* ---------- arranque ---------- */
  async function init() {
    try { const r = await fetch('admin/config.json?t=' + Date.now(), { cache: 'no-store' }); if (r.ok) CONFIG = await r.json(); } catch (e) { }
    if (!CONFIG) { $('vSetup').hidden = false; return; }
    const guardada = store.get('granja_pass');
    if (guardada) { try { TOKEN = await descifrar(CONFIG, guardada); return abrir(); } catch (e) { store.del('granja_pass'); } }
    $('vLogin').hidden = false; $('pass').focus();
  }
  $('fLogin').addEventListener('submit', async e => {
    e.preventDefault();
    const p = $('pass').value, btn = qs('button', e.target); btn.disabled = true;
    try { TOKEN = await descifrar(CONFIG, p); store.set('granja_pass', p, $('recordar').checked); $('vLogin').hidden = true; $('toast').className = 'toast'; abrir(); }
    catch (err) { toast('Contraseña incorrecta', true); }
    btn.disabled = false;
  });
  $('fSetup').addEventListener('submit', async e => {
    e.preventDefault();
    const tok = $('sToken').value.trim(), p = $('sPass').value, p2 = $('sPass2').value, btn = qs('button', e.target);
    if (p !== p2) return toast('Las contraseñas no coinciden', true);
    btn.disabled = true;
    try {
      const r = await gh('', {}, tok); const repo = await r.json();
      if (!r.ok || !repo.permissions || !repo.permissions.push) throw new Error('El token no tiene permiso de escritura en el repositorio');
      const cfg = await cifrar(tok, p);
      const ex = await gh('/contents/admin/config.json', {}, tok); const sha = ex.ok ? (await ex.json()).sha : undefined;
      const w = await gh('/contents/admin/config.json', { method: 'PUT', body: JSON.stringify({ message: 'Configurar pantalla de captura', content: utf8b64(JSON.stringify(cfg)), sha }) }, tok);
      if (!w.ok) throw new Error('No se pudo guardar la configuración');
      CONFIG = cfg; TOKEN = tok; store.set('granja_pass', p, false);
      $('sToken').value = ''; $('vSetup').hidden = true; toast('Configuración guardada'); abrir();
    } catch (err) { toast(err.message, true); }
    btn.disabled = false;
  });
  $('salir').addEventListener('click', () => { store.del('granja_pass'); TOKEN = null; location.reload(); });

  async function abrir() {
    $('vApp').hidden = false; $('whoBox').hidden = false;
    const q = store.get('granja_quien');
    if (q) { if ([...$('quien').options].some(o => o.value === q)) $('quien').value = q; else { $('quien').value = '__otro'; $('quienOtro').hidden = false; $('quienOtro').value = q; } }
    qsa('input[type=date]').forEach(i => { if (!i.value && i.name !== 'refuerzo' && i.name !== 'vence') i.value = hoy(); });
    try { DATA = await (await fetch('data.json?t=' + Date.now(), { cache: 'no-store' })).json(); } catch (e) { DATA = {}; }
    llenarListas(); tipoDinero(); refrescar(); setInterval(refrescar, 20000);
  }
  $('quien').addEventListener('change', () => { $('quienOtro').hidden = $('quien').value !== '__otro'; guardarQuien(); });
  $('quienOtro').addEventListener('input', guardarQuien);
  function guardarQuien() { store.set('granja_quien', quien(), true); }
  const quien = () => $('quien').value === '__otro' ? ($('quienOtro').value.trim() || 'Otra persona') : $('quien').value;

  /* ---------- pestañas ---------- */
  qsa('#tabs button').forEach(b => b.addEventListener('click', () => {
    qsa('#tabs button').forEach(x => x.setAttribute('aria-pressed', x === b));
    qsa('[data-panel]').forEach(p => p.hidden = p.dataset.panel !== b.dataset.t);
  }));
  qsa('.seg.sub').forEach(seg => qsa('button', seg).forEach(b => b.addEventListener('click', () => {
    qsa('button', seg).forEach(x => x.setAttribute('aria-pressed', x === b));
    qsa('[data-sp]', seg.parentElement).forEach(f => f.hidden = f.dataset.sp !== b.dataset.s);
  })));

  /* ---------- listas desde data.json ---------- */
  const TIPOS = { Sale: ['Concentrado', 'Sanidad y medicinas', 'Materiales y mantenimiento', 'Flete / comisión', 'Mano de obra', 'Animales', 'Servicios', 'Otro gasto'],
    Entra: ['Venta de lechones', 'Venta de cerdos', 'Aporte de socio', 'Otro ingreso'] };
  function tipoDinero() {
    const f = qs('form[data-tipo=dinero]'); const mv = qs('input[name=movimiento]:checked', f).value;
    const sel = f.tipo, prev = sel.value; sel.innerHTML = TIPOS[mv].map(t => `<option>${t}</option>`).join('');
    if (TIPOS[mv].includes(prev)) sel.value = prev;
    qs('[data-show=aporte]', f).hidden = sel.value !== 'Aporte de socio';
  }
  qsa('form[data-tipo=dinero] input[name=movimiento]').forEach(i => i.addEventListener('change', tipoDinero));
  qs('form[data-tipo=dinero] select[name=tipo]').addEventListener('change', tipoDinero);

  function opts(sel, items, vacio) {
    const prev = sel.value;
    sel.innerHTML = (vacio ? `<option value="">${vacio}</option>` : '') + items.map(i => typeof i === 'string' ? `<option>${esc(i)}</option>` : `<option value="${esc(i.v)}">${esc(i.t)}</option>`).join('');
    if (prev && [...sel.options].some(o => o.value === prev)) sel.value = prev;
  }
  function llenarListas() {
    const D = DATA || {};
    const hato = D.hato || [];
    const activas = hato.filter(h => /^activ/i.test(h.estado || 'Activa'));
    const cerdas = [...new Set([...activas.filter(h => !/verraco/i.test(h.tipo)).map(h => h.codigo), ...(D.celos || []).map(c => c.cerda), ...(D.partos || []).map(p => p.cerda)])].filter(Boolean);
    const verracos = [...new Set([...activas.filter(h => /verraco/i.test(h.tipo)).map(h => h.codigo), 'Inseminación'])];
    qsa('select[data-list=cerdas]').forEach(s => opts(s, cerdas));
    qsa('select[data-list=verracos]').forEach(s => opts(s, verracos, '—'));
    const camadas = (D.lechones || []).map(l => ({ c: l.cerdas[0], f: l.fecha, q: (l.vivos || 0) - (l.muertes || 0) - (l.ventas || 0) })).filter(l => l.q > 0);
    qsa('select[data-list=camadas]').forEach(s => opts(s, camadas.map(l => ({ v: l.c, t: `${l.c} · parto ${fmtD(l.f)} · quedan ${l.q}` }))));
    qsa('select[data-list=hatoActivos]').forEach(s => opts(s, activas.map(h => ({ v: h.codigo, t: `${h.codigo} · ${h.tipo}` }))));
    const lotes = [...new Set((D.engorde || []).map(m => m.lote))];
    qsa('select[data-list=lotes]').forEach(s => { opts(s, [...lotes, { v: '__nuevo', t: 'Lote nuevo…' }]); s.dispatchEvent(new Event('change')); });
    $('dlCerdas').innerHTML = [...cerdas, ...lotes].map(c => `<option value="${esc(c)}">`).join('');
    $('dlProductos').innerHTML = ['Farrowsure', 'E. coli (Porcilis Coliclos)', 'Porcilis PCV', 'Ferricol 200', 'Collzuril', 'Minoxel LA', 'Apetovit', 'Desparasitante', 'Vitaminas', ...(D.medicamentos || []).map(m => m.nombre)].map(c => `<option value="${esc(c)}">`).join('');
    $('dlMedicinas').innerHTML = (D.medicamentos || []).map(m => `<option value="${esc(m.nombre)}">`).join('');
    $('dlConcentrados').innerHTML = (D.concentrado || []).map(m => `<option value="${esc(m.tipo)}">`).join('');
    const precio = (D.costo && D.costo.precio) || 2300;
    const fv = qs('form[data-tipo=venta_lechones]'); if (!fv.precio.value) fv.precio.value = precio;
    hints();
  }
  function hints() {
    const D = DATA || {};
    const fm = qs('form[data-tipo=monta]'), fp = qs('form[data-tipo=parto]');
    const c1 = fm.cerda.value, pend = (D.celos || []).filter(c => c.cerda === c1);
    qs('[data-hint=monta]', fm).innerHTML = fm.fecha.value ? `Parto esperado: <b>${fmtD(addDays(fm.fecha.value, 114))}</b>` + (pend.length ? ` · Ojo: la ${esc(c1)} ya tiene una monta sin parto (${fmtD(pend[pend.length - 1].monta)}).` : '') : '';
    const c2 = fp.cerda.value, p2 = (D.celos || []).filter(c => c.cerda === c2);
    qs('[data-hint=parto]', fp).innerHTML = p2.length ? `Se anotará en la monta del <b>${fmtD(p2[p2.length - 1].monta)}</b> (parto esperado ${fmtD(p2[p2.length - 1].partoEsperado)}).` : `La ${esc(c2)} no tiene monta pendiente: se creará una fila nueva.`;
    const fmd = qs('form[data-tipo=medicina]'), m = (D.medicamentos || []).find(x => (x.nombre || '').toLowerCase() === fmd.nombre.value.trim().toLowerCase());
    qs('[data-hint=medicina]', fmd).innerHTML = m ? `Hay <b>${m.cantidad ?? 0}</b> frasco(s) según el Excel.` : (fmd.nombre.value ? 'Medicina nueva: se agregará a Bodega.' : '');
    const fc = qs('form[data-tipo=concentrado]'), k = (D.concentrado || []).find(x => (x.tipo || '').toLowerCase() === fc.tipo.value.trim().toLowerCase());
    qs('[data-hint=concentrado]', fc).innerHTML = k ? `Hay <b>${k.bolsas ?? 0}</b> bolsa(s) según el Excel.` : (fc.tipo.value ? 'Concentrado nuevo: se agregará a Bodega.' : '');
  }
  document.addEventListener('input', e => { if (e.target.closest('form[data-tipo]')) hints(); });
  document.addEventListener('change', e => { if (e.target.closest('form[data-tipo]')) hints(); });
  // venta de lechones: total automático
  (() => { const f = qs('form[data-tipo=venta_lechones]'); const calc = () => { if (f.cantidad.value && f.precio.value && !f.monto.dataset.touched) f.monto.value = f.cantidad.value * f.precio.value; };
    f.cantidad.addEventListener('input', calc); f.precio.addEventListener('input', calc); f.monto.addEventListener('input', () => f.monto.dataset.touched = 1); })();
  // hato baja: precio solo si vendida
  (() => { const f = qs('form[data-tipo=hato_baja]'); const s = () => qs('[data-show=ventaHato]', f).hidden = f.estado.value !== 'Vendida'; f.estado.addEventListener('change', s); s(); })();
  // engorde: lote nuevo y monto
  (() => { const f = qs('form[data-tipo=engorde]'); const s = () => { qs('[data-show=loteNuevo]', f).hidden = f.lote.value !== '__nuevo'; qs('[data-show=montoEng]', f).hidden = !['Venta', 'Compra'].includes(f.movimiento.value); };
    f.lote.addEventListener('change', s); f.movimiento.addEventListener('change', s); s(); })();
  // refuerzo rápido
  qsa('[data-ref]').forEach(a => a.addEventListener('click', e => { e.preventDefault(); const f = a.closest('form'); f.refuerzo.value = addDays(f.fecha.value || hoy(), +a.dataset.ref); }));

  /* ---------- enviar ---------- */
  function datos(f) {
    const d = {};
    new FormData(f).forEach((v, k) => { v = String(v).trim(); if (v !== '') d[k] = v; });
    qsa('input[type=checkbox]', f).forEach(c => d[c.name] = c.checked);
    return d;
  }
  function resumen(tipo, d) {
    switch (tipo) {
      case 'dinero': return `${d.movimiento === 'Entra' ? 'Entra' : 'Sale'} ${L(d.monto)} · ${d.tipo} · ${d.descripcion}`;
      case 'monta': return `Monta de la cerda ${d.cerda}${d.verraco ? ' con ' + d.verraco : ''} (${fmtD(d.fecha)})`;
      case 'parto': return `Parto de la ${d.cerda}: ${d.vivos} vivos (${fmtD(d.fecha)})`;
      case 'muerte_lechones': return `Murieron ${d.cantidad} lechón(es) de la ${d.cerda}`;
      case 'venta_lechones': return `Venta de ${d.cantidad} lechones de la ${d.cerda}${d.monto ? ' por ' + L(d.monto) : ''}`;
      case 'hato_alta': return `Entra ${d.tipo || 'cerda'} ${d.codigo}`;
      case 'hato_baja': return `${d.codigo}: ${d.estado}${d.monto ? ' por ' + L(d.monto) : ''}`;
      case 'engorde': return `Engorde ${d.lote}: ${d.movimiento} de ${d.cantidad}${d.monto ? ' por ' + L(d.monto) : ''}`;
      case 'vacuna': return `${d.producto} a ${d.cerda} (${fmtD(d.fecha)})`;
      case 'medicina': return `${d.nombre}: ${d.modo === 'sumar' ? '+' : '='}${d.cantidad} frascos`;
      case 'concentrado': return `${d.tipo}: ${d.modo === 'sumar' ? '+' : '='}${d.cantidad} bolsas`;
    }
    return tipo;
  }
  qsa('form[data-tipo]').forEach(f => f.addEventListener('submit', async e => {
    e.preventDefault();
    const tipo = f.dataset.tipo, d = datos(f);
    if (tipo === 'venta_lechones' && !d.conDinero) delete d.monto;
    if (tipo === 'engorde' && d.lote === '__nuevo') { if (!d.loteNuevo) return toast('Escribe el nombre del lote nuevo', true); d.lote = d.loteNuevo; }
    if (tipo === 'medicina' && d.vence) d.vence = d.vence + '-01';
    if (tipo === 'dinero' && d.tipo !== 'Aporte de socio') { delete d.socio; delete d.tipoAporte; }
    const btn = qs('button[type=submit]', f); btn.disabled = true;
    const ahora = new Date(), id = `${ahora.toISOString().replace(/[-:T]/g, '').slice(0, 14)}-${Math.random().toString(36).slice(2, 7)}`;
    const cap = { id, creado: ahora.toISOString(), quien: quien(), tipo, datos: d, resumen: resumen(tipo, d) };
    try {
      const r = await gh(`/contents/pendientes/${id}.json`, { method: 'PUT', body: JSON.stringify({ message: `Captura: ${cap.resumen}`.slice(0, 120), content: utf8b64(JSON.stringify(cap, null, 1)) }) });
      if (!r.ok) throw new Error('No se pudo guardar');
      toast('Guardado ✓ Se pasa al Excel en 1–3 minutos');
      const keep = ['fecha', 'movimiento', 'cerda', 'lote', 'precio', 'conDinero', 'modo'];
      qsa('input,select,textarea', f).forEach(i => { if (!keep.includes(i.name) && i.type !== 'radio' && i.type !== 'checkbox' && i.tagName !== 'SELECT') i.value = ''; });
      delete (f.monto || {}).dataset?.touched;
      refrescar();
    } catch (err) { toast('No se guardó: ' + err.message + '. Revisa tu conexión e inténtalo de nuevo.', true); }
    btn.disabled = false;
  }));

  /* ---------- listas de pendientes y aplicadas ---------- */
  async function refrescar() {
    if (!TOKEN) return;
    try {
      const r = await gh('/contents/pendientes');
      const arr = r.ok ? (await r.json()).filter(x => x.name.endsWith('.json')) : [];
      const items = await Promise.all(arr.map(async x => { try { const j = await (await gh('/contents/' + x.path, { headers: { Accept: 'application/vnd.github.raw+json' } })).json(); return { ...j, sha: x.sha, path: x.path }; } catch (e) { return { id: x.name, resumen: x.name, sha: x.sha, path: x.path }; } }));
      items.sort((a, b) => (a.creado || '') < (b.creado || '') ? 1 : -1);
      $('pend').innerHTML = items.length ? items.map(i => `<div class="item"><div>${esc(i.resumen)}<div class="m">${esc(i.quien || '')} · ${i.creado ? new Date(i.creado).toLocaleString('es-HN', { dateStyle: 'short', timeStyle: 'short' }) : ''}</div></div><button type="button" data-del="${esc(i.path)}" data-sha="${esc(i.sha)}">Borrar</button></div>`).join('') : '<p class="muted" style="font-size:13px">Nada pendiente.</p>';
    } catch (e) { $('pend').innerHTML = '<p class="muted" style="font-size:13px">No se pudo consultar.</p>'; }
    try {
      const r = await gh('/contents/capturas.json', { headers: { Accept: 'application/vnd.github.raw+json' } });
      const log = r.ok ? await r.json() : [];
      $('hechas').innerHTML = log.length ? log.slice(0, 15).map(i => `<div class="item"><div>${i.estado === 'ok' ? '✓' : '<b style="color:var(--bad)">✗</b>'} ${esc(i.resumen || i.tipo)}<div class="m">${esc(i.quien || '')} · ${new Date(i.aplicado).toLocaleString('es-HN', { dateStyle: 'short', timeStyle: 'short' })}${i.estado === 'ok' ? '' : ' · ' + esc(i.mensaje)}</div></div></div>`).join('') : '<p class="muted" style="font-size:13px">Todavía no hay capturas.</p>';
    } catch (e) { }
  }
  document.addEventListener('click', async e => {
    const b = e.target.closest('[data-del]'); if (!b) return;
    b.disabled = true;
    try { const r = await gh('/contents/' + b.dataset.del, { method: 'DELETE', body: JSON.stringify({ message: 'Borrar captura', sha: b.dataset.sha }) });
      if (!r.ok) throw new Error(); toast('Captura borrada'); }
    catch (err) { toast('Ya se estaba pasando al Excel o no se pudo borrar', true); }
    refrescar();
  });

  init();
})();
