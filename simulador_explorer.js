/* ==========================================================================
   Explorador de Datos — Simulador de riesgo para plagas y enfermedades nuevas
   Permite definir las condiciones climáticas (variables, umbrales y pesos) de un
   patógeno que no está en el SAT y ver, con el pronóstico vigente, qué municipios
   productores quedarían en cada nivel de riesgo.

   Datos: data/pronostico/pronostico_latest.json (el mismo del explorador del pronóstico)
          + geometría en data/geo/. Todo se calcula en el navegador.
   La lógica (variables del periodo, índice ponderado y semáforo) es la misma de los
   modelos en Python (Scripts/Modelos_Agroclimaticos_Python).
   Requiere Leaflet 1.9 (cargado en index.html).
   ========================================================================== */
(function () {
    'use strict';

    const RUTAS = {
        pronostico: 'data/pronostico/pronostico_latest.json',
        municipios: 'data/geo/municipios_bolivia.geojson',
        departamentos: 'data/geo/departamentos_bolivia.geojson',
        cultivos: 'data/geo/cultivos_municipios.json',
    };
    const MULTI = 'MULTIMODELO';
    const TODOS = 'TODOS';
    const TZ = 'America/La_Paz';
    const COLOR_OSCURO = '#232d39';
    const SIN_DATO = '#e5e7eb';
    const UMBRAL_LLUVIA_MM = 1;     // día de lluvia: P > 1 mm; día seco: P <= 1 mm (igual que en R y Python)
    const VENTANA_MIN = 3, VENTANA_MAX = 16;

    const CATEGORIAS = [
        { nombre: 'Normal', color: '#64a844' },
        { nombre: 'Alerta baja', color: '#fcd333' },
        { nombre: 'Alerta media', color: '#e88d2b' },
        { nombre: 'Alerta alta', color: '#c4161b' },
    ];

    // Variables del periodo que se pueden usar en una condición
    const VARS = [
        { k: 't', v: 't_media_periodo', nombre: 'Temperatura media', corto: 'T media', unidad: '°C', dec: 1, paso: 0.5,
          ayuda: 'Promedio de (Tmax + Tmin) / 2 en la ventana', def: { op: '>=', a: 20, b: 25 } },
        { k: 'tmax', v: 'tmax_media_periodo', nombre: 'Temperatura máxima media', corto: 'Tmax media', unidad: '°C', dec: 1, paso: 0.5,
          ayuda: 'Promedio de la temperatura máxima diaria', def: { op: '>=', a: 28, b: 32 } },
        { k: 'tmin', v: 'tmin_media_periodo', nombre: 'Temperatura mínima media', corto: 'Tmin media', unidad: '°C', dec: 1, paso: 0.5,
          ayuda: 'Promedio de la temperatura mínima diaria', def: { op: '<=', a: 5, b: 10 } },
        { k: 'precip', v: 'precip_acum_periodo', nombre: 'Precipitación acumulada', corto: 'Precipitación', unidad: 'mm', dec: 1, paso: 1,
          ayuda: 'Suma de la lluvia diaria en la ventana', def: { op: '>=', a: 30, b: 80 } },
        { k: 'hr', v: 'hr_media_periodo', nombre: 'Humedad relativa media', corto: 'HR media', unidad: '%', dec: 0, paso: 1,
          ayuda: 'Promedio de la humedad relativa diaria', def: { op: '>=', a: 80, b: 90 } },
        { k: 'lluvia', v: 'dias_lluvia_periodo', nombre: 'Días con lluvia', corto: 'Días con lluvia', unidad: 'días', dec: 0, paso: 1,
          ayuda: 'Días con más de 1 mm de lluvia', def: { op: '>=', a: 3, b: 6 } },
        { k: 'secos', v: 'dias_secos_periodo', nombre: 'Racha seca al final', corto: 'Racha seca', unidad: 'días', dec: 0, paso: 1,
          ayuda: 'Días seguidos con 1 mm o menos, contados hasta el último día', def: { op: '>=', a: 7, b: 10 } },
    ];
    const VAR = Object.fromEntries(VARS.map((x) => [x.k, x]));
    const OPS = { '>=': '≥', '>': '>', '<=': '≤', '<': '<', entre: 'entre' };
    const OPS_PY = { '>=': '>=', '>': '>', '<=': '<=', '<': '<' };

    // Los 15 indicadores del SAT, para partir de uno de ellos
    const c = (op, a, peso, b) => ({ on: true, op, a, b: b === undefined ? null : b, peso });
    const PLANTILLAS = {
        AJO: { cultivo: 'AJO', nombre: 'Fusariosis (pudrición basal)', ventana: 15, conds: { t: c('>=', 25, 1), secos: c('>=', 10, 1), precip: c('<', 20, 1) } },
        BANANO: { cultivo: 'BANANO', nombre: 'Sigatoka negra', ventana: 10, conds: { t: c('entre', 25, 1, 28), hr: c('>=', 90, 2), precip: c('>=', 50, 1) } },
        CAFE: { cultivo: 'CAFE', nombre: 'Broca del café', ventana: 14, conds: { t: c('entre', 20, 2, 30), precip: c('entre', 14, 1, 37.33) } },
        CEBOLLA: { cultivo: 'CEBOLLA', nombre: 'Mildiu (polvillo negro)', ventana: 10, conds: { t: c('entre', 6, 1, 22), hr: c('>=', 90, 2), lluvia: c('>=', 3, 1) } },
        DURAZNO_ARANUELA: { cultivo: 'DURAZNO', nombre: 'Arañuela roja', ventana: 10, conds: { t: c('>', 25, 2), secos: c('>=', 7, 1), hr: c('<', 50, 1) } },
        DURAZNO_OIDIO: { cultivo: 'DURAZNO', nombre: 'Oídio', ventana: 10, conds: { t: c('entre', 18, 1, 25), hr: c('entre', 50, 1, 70) } },
        FRUTILLA: { cultivo: 'FRUTILLA', nombre: 'Botrytis (moho gris)', ventana: 10, conds: { t: c('entre', 15, 1, 23), hr: c('>=', 90, 2), lluvia: c('>=', 2, 1) } },
        LECHUGA: { cultivo: 'LECHUGA', nombre: 'Sclerotinia sclerotiorum', ventana: 15, conds: { t: c('entre', 15, 1, 20), hr: c('>=', 85, 1), precip: c('>=', 60, 2) } },
        MAIZ_CHOCLO: { cultivo: 'MAIZ_CHOCLO', nombre: 'Gusano cogollero', ventana: 7, conds: { t: c('>=', 22, 2), precip: c('>=', 20, 1) } },
        MAIZ_EN_GRANO: { cultivo: 'MAIZ_EN_GRANO', nombre: 'Roya', ventana: 10, conds: { t: c('entre', 15, 1, 28), hr: c('>=', 75, 2), precip: c('>=', 30, 1) } },
        NARANJA: { cultivo: 'NARANJA', nombre: 'Mosca de la fruta', ventana: 15, conds: { t: c('entre', 20, 1, 30), secos: c('>=', 10, 2) } },
        PAPA: { cultivo: 'PAPA', nombre: 'Gorgojo de los Andes', ventana: 10, conds: { t: c('<=', 15, 1), precip: c('>=', 30, 2) } },
        QUINUA: { cultivo: 'QUINUA', nombre: 'Ticona', ventana: 10, conds: { t: c('entre', 15, 1, 25), precip: c('>=', 30, 2), lluvia: c('>=', 3, 1) } },
        TOMATE: { cultivo: 'TOMATE', nombre: 'Polilla (Tuta absoluta)', ventana: 10, conds: { t: c('>', 22, 2), secos: c('>=', 7, 1), hr: c('<', 60, 1) } },
        UVA: { cultivo: 'UVA', nombre: 'Mildiu de la uva', ventana: 7, conds: { t: c('>=', 10, 1), precip: c('>=', 10, 1), hr: c('>=', 85, 1) } },
    };

    // ======================================================================== núcleo de cálculo
    // (funciones puras; se exponen en window.SimuladorRiesgo para pruebas)

    const nulo = (x) => x === null || x === undefined || Number.isNaN(x);

    /** Variables del periodo para una fila (municipio) de un modelo: primeros n días.
     *  Un día faltante en la serie que usa una variable la deja en null (como en R/Python). */
    function variablesPeriodo(tmax, tmin, precip, hr, n) {
        const out = {};
        const tx = tmax.slice(0, n), tn = tmin.slice(0, n), p = precip.slice(0, n), h = hr.slice(0, n);
        const completa = (arr) => arr.length === n && arr.every((x) => !nulo(x));
        const media = (arr) => arr.reduce((s, x) => s + x, 0) / arr.length;
        const okT = completa(tx) && completa(tn);
        out.t_media_periodo = okT ? media(tx.map((x, i) => (x + tn[i]) / 2)) : null;
        out.tmax_media_periodo = completa(tx) ? media(tx) : null;
        out.tmin_media_periodo = completa(tn) ? media(tn) : null;
        const okP = completa(p);
        out.precip_acum_periodo = okP ? p.reduce((s, x) => s + x, 0) : null;
        out.dias_lluvia_periodo = okP ? p.filter((x) => x > UMBRAL_LLUVIA_MM).length : null;
        let racha = 0;
        p.forEach((x) => { racha = x <= UMBRAL_LLUVIA_MM ? racha + 1 : 0; });
        out.dias_secos_periodo = okP ? racha : null;
        out.hr_media_periodo = completa(h) ? media(h) : null;
        return out;
    }

    function cumple(valor, cond) {
        if (nulo(valor)) return null;
        switch (cond.op) {
            case '>=': return valor >= cond.a ? 1 : 0;
            case '>': return valor > cond.a ? 1 : 0;
            case '<=': return valor <= cond.a ? 1 : 0;
            case '<': return valor < cond.a ? 1 : 0;
            case 'entre': return valor >= Math.min(cond.a, cond.b) && valor <= Math.max(cond.a, cond.b) ? 1 : 0;
            default: return null;
        }
    }

    /** Índice ponderado y condiciones cumplidas. activas: [[k, cond], ...] */
    function evaluar(vars, activas) {
        const conds = {};
        let suma = 0, total = 0, faltante = false;
        activas.forEach(([k, cond]) => {
            const x = cumple(vars[VAR[k].v], cond);
            conds[k] = x;
            if (x === null) faltante = true;
            else suma += x * cond.peso;
            total += cond.peso;
        });
        if (faltante || total === 0) {
            Object.keys(conds).forEach((k) => { conds[k] = null; });
            return { indice: null, alerta: -1, conds };
        }
        const indice = suma / total;
        return { indice, alerta: clasificar(indice), conds };
    }

    function clasificar(indice) {
        if (nulo(indice)) return -1;
        const i = Math.round(indice * 1e9) / 1e9;
        return i === 0 ? 0 : i < 0.5 ? 1 : i < 1 ? 2 : 3;
    }

    window.SimuladorRiesgo = { variablesPeriodo, evaluar, clasificar, cumple, PLANTILLAS, VARS };

    // ======================================================================== interfaz
    const st = { nombre: '', cultivo: 'PAPA', ventana: 10, modelo: MULTI, plantilla: '', conds: {}, gid: null };
    let P = null, MUN = null, DEP = null, CULT = null;
    let propsDe = {}, filaDe = {};
    let gids = [];                 // municipios evaluados
    let agg = {};                  // modelo -> [vars por municipio] (cache por cultivo y ventana)
    let res = {};                  // modelo -> [resultado por municipio]
    let mapa = null, capa = null, iniciado = false;

    const $ = (id) => document.getElementById(id);
    const el = (tag, cls, texto) => {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        if (texto !== undefined) e.textContent = texto;
        return e;
    };
    const num = (v, dec = 1) => v.toLocaleString('es-BO', { maximumFractionDigits: dec, minimumFractionDigits: 0 });
    const idx = (v) => v.toLocaleString('es-BO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const aFecha = (iso) => new Date(iso + 'T12:00:00Z');
    const fmtCorto = new Intl.DateTimeFormat('es-BO', { day: 'numeric', month: 'short', timeZone: 'UTC' });
    const fmtLargo = new Intl.DateTimeFormat('es-BO', { weekday: 'long', day: 'numeric', month: 'long', timeZone: 'UTC' });
    const corto = (iso) => fmtCorto.format(aFecha(iso)).replace('.', '');
    const hoyBolivia = () => new Intl.DateTimeFormat('en-CA', { timeZone: TZ }).format(new Date());
    const etiquetaModelo = (m) => (m === MULTI ? 'Multimodelo' : m);
    const colorCat = (c) => (c >= 0 ? CATEGORIAS[c].color : SIN_DATO);
    const nombreCat = (c) => (c >= 0 ? CATEGORIAS[c].nombre : 'Sin dato');
    const etiquetaCultivo = (k) => (k === TODOS ? 'Todos los cultivos' : CULT[k] ? CULT[k].etiqueta : k);
    const infoMun = (gid) => (st.cultivo !== TODOS && CULT[st.cultivo] ? CULT[st.cultivo].municipios[gid] : null);
    const nombreMun = (gid) => {
        const i = infoMun(gid);
        return i ? i.nombres.join(' / ') : (propsDe[gid] ? propsDe[gid].municipio : gid);
    };
    const activas = () => VARS.filter((x) => st.conds[x.k] && st.conds[x.k].on).map((x) => [x.k, st.conds[x.k]]);
    const fmtVal = (k, v) => (nulo(v) ? '—' : `${num(v, VAR[k].dec)} ${VAR[k].unidad}`);

    function textoCond(k, cond, py) {
        const x = VAR[k];
        const n = (v) => (py ? String(v) : num(v, 2));
        const u = x.unidad === 'días' ? ' días' : ` ${x.unidad}`;
        if (cond.op === 'entre') return `${x.corto} entre ${n(Math.min(cond.a, cond.b))} y ${n(Math.max(cond.a, cond.b))}${u}`;
        return `${x.corto} ${py ? OPS_PY[cond.op] : OPS[cond.op]} ${n(cond.a)}${u}`;
    }

    // ---------------------------------------------------------------- carga
    async function cargarJSON(url, sinCache) {
        const r = await fetch(url, sinCache ? { cache: 'no-cache' } : {});
        if (!r.ok) throw new Error(`${url} respondió ${r.status}`);
        return r.json();
    }

    async function iniciar() {
        if (iniciado) return;
        iniciado = true;
        try {
            if (typeof L === 'undefined') throw new Error('no se cargó la librería de mapas (Leaflet)');
            [P, MUN, DEP, CULT] = await Promise.all([
                cargarJSON(RUTAS.pronostico, true), cargarJSON(RUTAS.municipios),
                cargarJSON(RUTAS.departamentos), cargarJSON(RUTAS.cultivos),
            ]);
        } catch (err) {
            iniciado = false;
            $('sm-cargando').textContent = `No se pudo cargar el pronóstico: ${err.message}. Recarga la página en unos minutos.`;
            return;
        }
        MUN.features.forEach((f) => { propsDe[f.properties.GID_3] = f.properties; });
        P.municipios.forEach((g, i) => { filaDe[g] = i; });

        $('sm-cargando').hidden = true;
        $('sm-contenido').hidden = false;
        const desdeEnlace = leerEnlace();
        if (!desdeEnlace) aplicarPlantilla('', { cultivo: 'PAPA' });
        construirControles();
        mostrarCorrida();
        crearMapa();
        recalcularTodo({ recrearCapa: true });
    }

    // ---------------------------------------------------------------- configuración
    function condicionesPorDefecto() {
        const conds = {};
        VARS.forEach((x) => { conds[x.k] = { on: false, op: x.def.op, a: x.def.a, b: x.def.b, peso: 1 }; });
        return conds;
    }

    function aplicarPlantilla(clave, extra = {}) {
        const conds = condicionesPorDefecto();
        if (clave && PLANTILLAS[clave]) {
            const p = PLANTILLAS[clave];
            Object.entries(p.conds).forEach(([k, v]) => {
                conds[k] = { ...v, b: v.b === null ? VAR[k].def.b : v.b };
            });
            st.cultivo = p.cultivo;
            st.ventana = p.ventana;
            st.nombre = `${p.nombre} (variante)`;
        } else {
            conds.t.on = true;
            conds.precip.on = true;
            st.nombre = '';
            if (extra.cultivo) st.cultivo = extra.cultivo;
        }
        st.plantilla = clave;
        st.conds = conds;
        st.gid = null;
    }

    function config() {
        const conds = {};
        VARS.forEach((x) => {
            const c = st.conds[x.k];
            if (c.on) conds[x.k] = { op: c.op, a: c.a, ...(c.op === 'entre' ? { b: c.b } : {}), peso: c.peso };
        });
        return { v: 1, nombre: st.nombre, cultivo: st.cultivo, ventana: st.ventana, modelo: st.modelo, condiciones: conds };
    }

    function cargarConfig(cfg) {
        const conds = condicionesPorDefecto();
        Object.entries(cfg.condiciones || {}).forEach(([k, c]) => {
            if (!VAR[k] || !OPS[c.op]) return;
            conds[k] = { on: true, op: c.op, a: +c.a, b: c.b !== undefined ? +c.b : VAR[k].def.b, peso: Math.min(3, Math.max(1, +c.peso || 1)) };
        });
        st.conds = conds;
        st.nombre = String(cfg.nombre || '').slice(0, 80);
        st.cultivo = cfg.cultivo === TODOS || CULT[cfg.cultivo] ? cfg.cultivo : 'PAPA';
        st.ventana = Math.min(VENTANA_MAX, Math.max(VENTANA_MIN, +cfg.ventana || 10));
        st.modelo = cfg.modelo || MULTI;
        st.plantilla = '';
        st.gid = null;
    }

    // Enlace compartible: ?sim=<config en base64url>
    function codificar(obj) {
        const bytes = new TextEncoder().encode(JSON.stringify(obj));
        let bin = '';
        bytes.forEach((b) => { bin += String.fromCharCode(b); });
        return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    }
    function decodificar(txt) {
        const b64 = txt.replace(/-/g, '+').replace(/_/g, '/');
        const bin = atob(b64 + '==='.slice((b64.length + 3) % 4));
        return JSON.parse(new TextDecoder().decode(Uint8Array.from(bin, (ch) => ch.charCodeAt(0))));
    }
    function leerEnlace() {
        const s = new URLSearchParams(location.search).get('sim');
        if (!s) return false;
        try { cargarConfig(decodificar(s)); return true; } catch (e) { return false; }
    }
    function enlace() {
        const u = new URL(location.href);
        u.search = '';
        u.searchParams.set('sim', codificar(config()));
        u.hash = 'explorar';
        return u.toString();
    }

    // ---------------------------------------------------------------- controles de la barra lateral
    function construirControles() {
        const sc = $('sm-cultivo');
        sc.innerHTML = '';
        Object.keys(CULT).sort((a, b) => etiquetaCultivo(a).localeCompare(etiquetaCultivo(b), 'es'))
            .forEach((k) => sc.add(new Option(`${etiquetaCultivo(k)} (${Object.keys(CULT[k].municipios).length} municipios)`, k)));
        sc.add(new Option(`Todos los municipios productores (${P.municipios.length})`, TODOS));

        const sp = $('sm-plantilla');
        sp.innerHTML = '';
        sp.add(new Option('En blanco', ''));
        const grupo = document.createElement('optgroup');
        grupo.label = 'Indicadores del SAT';
        Object.entries(PLANTILLAS)
            .sort((a, b) => etiquetaCultivo(a[1].cultivo).localeCompare(etiquetaCultivo(b[1].cultivo), 'es'))
            .forEach(([k, p]) => grupo.append(new Option(`${etiquetaCultivo(p.cultivo)} – ${p.nombre}`, k)));
        sp.append(grupo);

        const rv = $('sm-ventana');
        rv.min = VENTANA_MIN;
        rv.max = Math.min(VENTANA_MAX, P.fechas.length);

        sincronizarControles();

        $('sm-nombre').addEventListener('input', () => { st.nombre = $('sm-nombre').value; actualizarEncabezado(); dibujarFormula(); });
        sc.addEventListener('change', () => { st.cultivo = sc.value; st.gid = null; recalcularTodo({ recrearCapa: true }); });
        sp.addEventListener('change', () => {
            aplicarPlantilla(sp.value, { cultivo: st.cultivo });
            sincronizarControles();
            recalcularTodo({ recrearCapa: true });
        });
        rv.addEventListener('input', () => { st.ventana = +rv.value; $('sm-ventana-txt').textContent = `${st.ventana} días`; recalcularTodo({}); });
        $('sm-modelo').addEventListener('change', () => { st.modelo = $('sm-modelo').value; evaluarTodo(); });
        $('sm-enlace').addEventListener('click', copiarEnlace);
        $('sm-csv').addEventListener('click', descargarCSV);
        $('sm-py').addEventListener('click', descargarPython);
        $('sm-json').addEventListener('click', descargarJSON);
        $('sm-json-abrir').addEventListener('click', () => $('sm-json-archivo').click());
        $('sm-json-archivo').addEventListener('change', abrirJSON);
        $('sm-reiniciar').addEventListener('click', () => {
            aplicarPlantilla('', { cultivo: st.cultivo });
            sincronizarControles();
            recalcularTodo({});
        });
    }

    function sincronizarControles() {
        $('sm-nombre').value = st.nombre;
        $('sm-cultivo').value = st.cultivo;
        $('sm-plantilla').value = st.plantilla;
        $('sm-ventana').value = st.ventana;
        $('sm-ventana-txt').textContent = `${st.ventana} días`;
        construirTarjetas();
    }

    function poblarModelos() {
        const sm = $('sm-modelo');
        sm.innerHTML = '';
        P.modelos.slice().sort((a, b) => (a === MULTI ? -1 : b === MULTI ? 1 : 0)).forEach((m) => {
            const ok = cubre(m);
            const o = new Option(etiquetaModelo(m) + (ok ? '' : ' (no cubre la ventana)'), m);
            o.disabled = !ok;
            sm.add(o);
        });
        if (!cubre(st.modelo)) st.modelo = MULTI;
        sm.value = st.modelo;
    }

    // ---------------------------------------------------------------- tarjetas de variables
    function construirTarjetas() {
        const cont = $('sm-vars');
        cont.innerHTML = '';
        VARS.forEach((x) => {
            const cond = st.conds[x.k];
            const card = el('div', 'sm-var' + (cond.on ? ' is-on' : ''));
            card.id = `sm-var-${x.k}`;

            const head = el('label', 'sm-var-head');
            const chk = el('input');
            chk.type = 'checkbox';
            chk.checked = cond.on;
            const tit = el('span', 'sm-var-tit');
            tit.append(el('span', 'sm-var-name', x.nombre), el('span', 'sm-var-help', x.ayuda));
            head.append(chk, tit, el('span', 'sm-var-unit', x.unidad));
            card.append(head);

            const body = el('div', 'sm-var-body');
            const op = el('select', 'sm-op');
            op.setAttribute('aria-label', `Condición para ${x.nombre}`);
            Object.entries(OPS).forEach(([k, t]) => op.add(new Option(t, k)));
            op.value = cond.op;
            const a = el('input', 'sm-num');
            a.type = 'number'; a.step = x.paso; a.value = cond.a;
            a.setAttribute('aria-label', `Umbral de ${x.nombre}`);
            const y = el('span', 'sm-y', 'y');
            const b = el('input', 'sm-num');
            b.type = 'number'; b.step = x.paso; b.value = cond.b;
            b.setAttribute('aria-label', `Umbral superior de ${x.nombre}`);
            const peso = el('div', 'sm-peso');
            peso.setAttribute('role', 'group');
            peso.setAttribute('aria-label', 'Peso de la condición');
            peso.append(el('span', 'sm-peso-lbl', 'Peso'));
            [1, 2, 3].forEach((p) => {
                const pb = el('button', 'sm-peso-btn' + (cond.peso === p ? ' is-sel' : ''), String(p));
                pb.type = 'button';
                pb.setAttribute('aria-pressed', String(cond.peso === p));
                pb.addEventListener('click', () => {
                    cond.peso = p;
                    peso.querySelectorAll('.sm-peso-btn').forEach((q) => {
                        const s = q.textContent === String(p);
                        q.classList.toggle('is-sel', s);
                        q.setAttribute('aria-pressed', String(s));
                    });
                    alCambiarCondicion();
                });
                peso.append(pb);
            });
            const umbral = el('div', 'sm-umbral');
            umbral.append(op, a, y, b);
            body.append(umbral, peso);
            card.append(body);

            const hist = el('div', 'sm-hist');
            hist.id = `sm-hist-${x.k}`;
            card.append(hist);
            card.append(el('p', 'sm-var-stat'));

            const mostrarB = () => { y.hidden = b.hidden = op.value !== 'entre'; };
            mostrarB();
            const habilitar = () => {
                [op, a, b].forEach((e) => { e.disabled = !cond.on; });
                peso.querySelectorAll('button').forEach((q) => { q.disabled = !cond.on; });
                card.classList.toggle('is-on', cond.on);
            };
            habilitar();
            chk.addEventListener('change', () => { cond.on = chk.checked; habilitar(); alCambiarCondicion(); });
            op.addEventListener('change', () => { cond.op = op.value; mostrarB(); alCambiarCondicion(); });
            a.addEventListener('input', () => { if (a.value !== '' && !isNaN(+a.value)) { cond.a = +a.value; alCambiarCondicion(); } });
            b.addEventListener('input', () => { if (b.value !== '' && !isNaN(+b.value)) { cond.b = +b.value; alCambiarCondicion(); } });
            cont.append(card);
        });
    }

    function alCambiarCondicion() {
        if (st.plantilla) { st.plantilla = ''; $('sm-plantilla').value = ''; }
        evaluarTodo();
    }

    // Histograma de la variable en los municipios evaluados (verde = cumple la condición)
    function dibujarHistograma(x) {
        const cont = $(`sm-hist-${x.k}`);
        if (!cont) return;
        const card = $(`sm-var-${x.k}`), stat = card.querySelector('.sm-var-stat');
        const cond = st.conds[x.k];
        const vals = (agg[st.modelo] || []).map((v) => v[x.v]).filter((v) => !nulo(v));
        cont.innerHTML = '';
        if (!vals.length) { stat.textContent = 'Sin datos para este modelo y ventana.'; return; }
        let lo = Math.min(...vals), hi = Math.max(...vals);
        const umbrales = cond.op === 'entre' ? [cond.a, cond.b] : [cond.a];
        if (x.unidad === 'días') { lo = Math.min(lo, 0); hi = Math.max(hi, 1); }
        if (hi - lo < 1e-9) { lo -= 1; hi += 1; }
        const nb = x.unidad === 'días' ? Math.min(17, Math.round(hi - lo) + 1) : 16;
        const ancho = (hi - lo) / (x.unidad === 'días' ? nb - 1 : nb);
        const bin = (v) => {
            const i = x.unidad === 'días' ? Math.round((v - lo) / ancho) : Math.floor((v - lo) / ancho);
            return Math.max(0, Math.min(nb - 1, i));
        };
        const si = new Array(nb).fill(0), no = new Array(nb).fill(0);
        vals.forEach((v) => { if (cond.on && cumple(v, cond) === 1) si[bin(v)] += 1; else no[bin(v)] += 1; });
        const maxN = Math.max(...si.map((s, i) => s + no[i]));
        const W = 280, H = 52, bw = W / nb;
        let svg = `<svg viewBox="0 0 ${W} ${H + 2}" preserveAspectRatio="none" role="img" aria-label="Distribución de ${x.nombre} en los municipios">`;
        for (let i = 0; i < nb; i++) {
            const hs = (si[i] / maxN) * H, hn = (no[i] / maxN) * H;
            if (hn) svg += `<rect x="${i * bw + 1}" y="${H - hn - hs}" width="${bw - 2}" height="${hn}" fill="${cond.on ? '#cfd6de' : '#dfe4ea'}"/>`;
            if (hs) svg += `<rect x="${i * bw + 1}" y="${H - hs}" width="${bw - 2}" height="${hs}" fill="#e88d2b"/>`;
        }
        if (cond.on) {
            const xr = (v) => {
                const pos = x.unidad === 'días' ? (v - lo) / ancho * bw + bw / 2 : (v - lo) / (hi - lo) * W;
                return Math.max(1, Math.min(W - 1, pos));
            };
            umbrales.forEach((u) => {
                if (u >= lo - ancho && u <= hi + ancho) {
                    svg += `<line x1="${xr(u)}" x2="${xr(u)}" y1="0" y2="${H}" stroke="#232d39" stroke-width="1.5" stroke-dasharray="3 2" vector-effect="non-scaling-stroke"/>`;
                }
            });
        }
        svg += `<line x1="0" x2="${W}" y1="${H + 1}" y2="${H + 1}" stroke="#b8c1cc" stroke-width="1" vector-effect="non-scaling-stroke"/></svg>`;
        cont.innerHTML = svg;
        const ejes = el('div', 'sm-hist-ejes');
        ejes.append(el('span', '', num(lo, x.dec)), el('span', '', num(hi, x.dec)));
        cont.append(ejes);
        const n = vals.length;
        if (cond.on) {
            const nSi = si.reduce((s, v) => s + v, 0);
            stat.innerHTML = '';
            stat.append(el('b', '', `Se cumple en ${nSi} de ${n}`), ` municipios · ${textoCond(x.k, cond)}`);
        } else {
            stat.textContent = `No se usa en el cálculo · rango ${num(Math.min(...vals), x.dec)} a ${num(Math.max(...vals), x.dec)} ${x.unidad}`;
        }
    }

    // ---------------------------------------------------------------- cálculo
    function cubre(m) {
        return (agg[m] || []).some((v) => !nulo(v.t_media_periodo) || !nulo(v.precip_acum_periodo));
    }

    function agregar() {
        gids = st.cultivo === TODOS
            ? P.municipios.slice()
            : Object.keys(CULT[st.cultivo].municipios).filter((g) => filaDe[g] !== undefined);
        const D = P.datos, n = st.ventana;
        agg = {};
        P.modelos.forEach((m) => {
            agg[m] = gids.map((g) => {
                const f = filaDe[g];
                const s = (v) => (D[v] && D[v][m] ? D[v][m][f] : []);
                return variablesPeriodo(s('temperature_2m_max'), s('temperature_2m_min'),
                    s('precipitation_sum'), s('relative_humidity_2m'), n);
            });
        });
    }

    function calcular() {
        const act = activas();
        res = {};
        P.modelos.forEach((m) => { res[m] = agg[m].map((v) => evaluar(v, act)); });
    }

    function recalcularTodo({ recrearCapa: recrear }) {
        agregar();
        poblarModelos();
        if (st.gid && !gids.includes(st.gid)) st.gid = null;
        if (recrear || !capa) recrearCapa();
        evaluarTodo();
    }

    function evaluarTodo() {
        calcular();
        capa.setStyle(estilo);
        VARS.forEach(dibujarHistograma);
        actualizarEncabezado();
        dibujarFormula();
        dibujarResumen();
        dibujarRanking();
        dibujarDetalle();
        const sin = P.modelos.filter((m) => m !== MULTI && !cubre(m));
        $('sm-nota-modelos').textContent = sin.length
            ? `${sin.join(', ')} no ${sin.length > 1 ? 'cubren' : 'cubre'} los ${st.ventana} días de la ventana, por eso no ${sin.length > 1 ? 'tienen' : 'tiene'} resultado.`
            : '';
        try { history.replaceState(null, '', location.pathname + location.hash); } catch (e) { /* sin historial */ }
    }

    // ---------------------------------------------------------------- mapa
    function crearMapa() {
        mapa = L.map('sm-mapa', { zoomSnap: 0.25, scrollWheelZoom: false, attributionControl: false });
        L.geoJSON(MUN, { interactive: false, style: { color: '#d5dbe2', weight: 0.5, fillColor: '#f1f3f5', fillOpacity: 1 } }).addTo(mapa);
        mapa.createPane('sm-deptos').style.zIndex = 450;
        mapa.createPane('sm-etiquetas').style.zIndex = 460;
        mapa.getPane('sm-deptos').style.pointerEvents = 'none';
        mapa.getPane('sm-etiquetas').style.pointerEvents = 'none';
        const deptos = L.geoJSON(DEP, { pane: 'sm-deptos', interactive: false, style: { color: '#6b7785', weight: 1.2, fill: false } }).addTo(mapa);
        DEP.features.forEach((f) => {
            L.tooltip({ permanent: true, direction: 'center', className: 'pr-depto-label', pane: 'sm-etiquetas' })
                .setLatLng(L.geoJSON(f).getBounds().getCenter()).setContent(f.properties.departamento).addTo(mapa);
        });
        mapa.fitBounds(deptos.getBounds(), { padding: [10, 10] });
        dibujarLeyenda();
    }

    const pos = (gid) => gids.indexOf(gid);
    function dato(gid, modelo = st.modelo) {
        const i = pos(gid);
        if (i < 0 || !res[modelo]) return null;
        return { ...res[modelo][i], vars: agg[modelo][i], i };
    }

    function estilo(feature) {
        const gid = feature.properties.GID_3;
        const d = dato(gid);
        const sel = gid === st.gid;
        const base = { color: sel ? COLOR_OSCURO : '#6b7785', weight: sel ? 3 : 0.6, fillOpacity: 1 };
        if (!d || d.alerta < 0) return { ...base, fillColor: SIN_DATO, dashArray: '3' };
        return { ...base, fillColor: colorCat(d.alerta), dashArray: null };
    }

    function tooltip(gid) {
        const d = dato(gid);
        const div = el('div');
        div.append(el('b', '', nombreMun(gid)), el('br'), propsDe[gid] ? propsDe[gid].departamento : '', el('br'));
        if (!d || d.alerta < 0) div.append(activas().length ? 'Sin dato para este modelo' : 'Activa al menos una variable');
        else div.append(`${nombreCat(d.alerta)} (índice ${idx(d.indice)})`);
        return div;
    }

    function recrearCapa() {
        if (capa) mapa.removeLayer(capa);
        const set = new Set(gids);
        capa = L.geoJSON(
            { type: 'FeatureCollection', features: MUN.features.filter((f) => set.has(f.properties.GID_3)) },
            {
                style: estilo,
                onEachFeature: (f, c) => {
                    c.bindTooltip(() => tooltip(f.properties.GID_3), { sticky: true, className: 'pr-tooltip' });
                    c.on('click', () => seleccionar(f.properties.GID_3));
                },
            },
        ).addTo(mapa);
    }

    function seleccionar(gid) {
        st.gid = gid;
        capa.setStyle(estilo);
        capa.eachLayer((c) => { if (c.feature.properties.GID_3 === gid) c.bringToFront(); });
        dibujarRanking();
        dibujarDetalle();
    }

    function dibujarLeyenda() {
        const cont = $('sm-leyenda');
        cont.innerHTML = '';
        cont.append(el('span', 'pr-legend-title', 'Nivel de riesgo'));
        CATEGORIAS.forEach((c) => {
            const it = el('div', 'pr-legend-item al-legend-item');
            const sw = el('span', 'pr-legend-swatch');
            sw.style.background = c.color;
            it.append(sw, el('span', '', c.nombre));
            cont.append(it);
        });
        const nd = el('div', 'pr-legend-item pr-legend-nodata al-legend-item');
        nd.append(el('span', 'pr-legend-swatch'), el('span', '', 'sin dato'));
        cont.append(nd);
    }

    // ---------------------------------------------------------------- encabezado, fórmula y resumen
    function mostrarCorrida() {
        const hora = new Intl.DateTimeFormat('es-BO', {
            timeZone: TZ, day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
        }).format(new Date(P.generado_utc));
        $('sm-corrida').innerHTML = '';
        $('sm-corrida').append('Pronóstico del ', el('strong', '', hora), ' (hora de Bolivia).');
        const atraso = Math.round((aFecha(hoyBolivia()) - aFecha(P.fecha_inicio)) / 86400000);
        if (atraso >= 2) {
            $('sm-alerta').textContent = `El simulador usa el pronóstico del ${fmtLargo.format(aFecha(P.fecha_inicio))}, ` +
                `que no se ha actualizado en ${atraso} días.`;
            $('sm-alerta').hidden = false;
        }
    }

    function actualizarEncabezado() {
        const nom = st.nombre.trim() || 'Plaga o enfermedad sin nombre';
        $('sm-titulo').textContent = `${nom} en ${etiquetaCultivo(st.cultivo).toLowerCase()}`;
        $('sm-subtitulo').textContent = `${gids.length} municipios · ${etiquetaModelo(st.modelo).toLowerCase()} · simulación`;
        const f = $('sm-fecha');
        f.innerHTML = '';
        const n = Math.min(st.ventana, P.fechas.length);
        f.append(`${corto(P.fechas[0])} – ${corto(P.fechas[n - 1])}`, el('small', '', `ventana de ${n} días`));
    }

    function dibujarFormula() {
        const cont = $('sm-formula'), act = activas();
        cont.innerHTML = '';
        if (!act.length) {
            cont.append(el('p', 'sm-formula-vacia', 'Activa al menos una variable para calcular el riesgo.'));
            return;
        }
        const total = act.reduce((s, [, c]) => s + c.peso, 0);
        const p = el('p', 'sm-formula-txt');
        p.append(el('b', '', 'Alerta si: '), act.map(([k, c]) => textoCond(k, c)).join(' y ') + '.');
        const ul = el('ul', 'sm-formula-list');
        act.forEach(([k, c]) => {
            const li = el('li');
            const w = el('span', 'sm-formula-w', `${c.peso}/${total}`);
            w.title = `Aporta ${num(100 * c.peso / total, 0)} % del índice`;
            li.append(el('span', '', textoCond(k, c)), w);
            ul.append(li);
        });
        cont.append(p, ul, el('p', 'al-nota',
            `Índice = suma de los pesos de las condiciones que se cumplen ÷ ${total}. ` +
            '0 = normal; menos de 0,5 = alerta baja; menos de 1 = alerta media; 1 = alerta alta.'));
    }

    function dibujarResumen() {
        const r = res[st.modelo] || [];
        const ha = (i) => { const m = infoMun(gids[i]); return m && m.sup_ha ? m.sup_ha : 0; };
        const totalHa = gids.reduce((s, _, i) => s + ha(i), 0);
        const cont = $('sm-resumen');
        cont.innerHTML = '';
        [3, 2, 1, 0, -1].forEach((cat) => {
            const ix = r.map((x, i) => (x.alerta === cat ? i : -1)).filter((i) => i >= 0);
            if (cat === -1 && !ix.length) return;
            const sup = ix.reduce((s, i) => s + ha(i), 0);
            const t = el('div', 'al-tile' + (cat < 0 ? ' al-tile-nd' : ''));
            t.style.setProperty('--c', colorCat(cat));
            t.append(el('span', 'al-tile-n', String(ix.length)), el('span', 'al-tile-cat', nombreCat(cat)));
            if (totalHa > 0 && cat >= 0) t.append(el('span', 'al-tile-ha', `${num(100 * sup / totalHa, 0)} % de la superficie`));
            cont.append(t);
        });
    }

    // ---------------------------------------------------------------- ranking y detalle
    function dibujarRanking() {
        const cont = $('sm-ranking');
        cont.innerHTML = '';
        const filas = gids.map((g) => ({ g, d: dato(g), ha: (infoMun(g) || {}).sup_ha || 0 }))
            .filter((x) => x.d && x.d.alerta > 0)
            .sort((a, b) => b.d.indice - a.d.indice || b.ha - a.ha);
        if (!filas.length) {
            cont.append(el('p', 'al-vacio', activas().length ? 'Ningún municipio en riesgo con estas condiciones.' : 'Activa al menos una variable.'));
            return;
        }
        const ol = el('ol', 'al-rank-list');
        filas.slice(0, 10).forEach(({ g, d, ha }) => {
            const b = el('button', 'al-rank-item' + (g === st.gid ? ' is-sel' : ''));
            b.type = 'button';
            const dot = el('span', 'al-dot');
            dot.style.background = colorCat(d.alerta);
            const txt = el('span', 'al-rank-txt');
            txt.append(el('span', 'al-rank-nom', nombreMun(g)),
                el('span', 'al-rank-meta', `${propsDe[g] ? propsDe[g].departamento : ''}${ha ? ' · ' + num(ha, 0) + ' ha' : ''}`));
            b.append(dot, txt, el('span', 'al-rank-idx', idx(d.indice)));
            b.addEventListener('click', () => seleccionar(g));
            const li = el('li');
            li.append(b);
            ol.append(li);
        });
        cont.append(ol);
        if (filas.length > 10) cont.append(el('p', 'al-nota', `y ${filas.length - 10} municipios más en riesgo.`));
    }

    function dibujarDetalle() {
        const vacio = $('sm-detalle-vacio'), cont = $('sm-detalle');
        if (!st.gid || pos(st.gid) < 0) {
            st.gid = null;
            vacio.hidden = false;
            cont.hidden = true;
            return;
        }
        vacio.hidden = true;
        cont.hidden = false;
        cont.innerHTML = '';
        const g = st.gid, p = propsDe[g] || {}, info = infoMun(g) || {}, d = dato(g);

        const cab = el('div', 'al-det-head');
        const tit = el('div');
        tit.append(el('h4', '', nombreMun(g)));
        let meta = `${p.departamento || ''}${p.provincia ? ', provincia ' + p.provincia : ''}.`;
        if (info.sup_ha) meta += ` ${num(info.sup_ha, 0)} ha de ${etiquetaCultivo(st.cultivo).toLowerCase()} (censo agropecuario).`;
        tit.append(el('p', 'pr-series-meta', meta));
        const badge = el('span', 'al-badge', d && d.alerta >= 0 ? `${nombreCat(d.alerta)} · índice ${idx(d.indice)}` : 'Sin dato');
        badge.style.setProperty('--c', colorCat(d ? d.alerta : -1));
        cab.append(tit, badge);
        cont.append(cab);

        cont.append(el('h5', '', `Valores en la ventana · ${etiquetaModelo(st.modelo)}`));
        const tabla = el('table', 'al-cond-table');
        const thead = el('tr');
        ['Variable', 'Pronóstico', 'Condición', 'Peso', ''].forEach((h) => thead.append(el('th', '', h)));
        tabla.append(thead);
        VARS.forEach((x) => {
            const cond = st.conds[x.k];
            const valor = d ? d.vars[x.v] : null;
            const tr = el('tr', cond.on ? '' : 'sm-fila-off');
            tr.append(el('td', '', x.corto), el('td', 'al-num', fmtVal(x.k, valor)),
                el('td', '', cond.on ? textoCond(x.k, cond).replace(x.corto + ' ', '') : 'no se usa'),
                el('td', 'al-num', cond.on ? String(cond.peso) : ''));
            const k = cond.on ? cumple(valor, cond) : undefined;
            tr.append(el('td', 'al-estado ' + (k === 1 ? 'si' : k === 0 ? 'no' : ''),
                k === 1 ? '✓ se cumple' : k === 0 ? '✗ no se cumple' : cond.on ? 'sin dato' : ''));
            tabla.append(tr);
        });
        cont.append(tabla);

        cont.append(el('h5', '', 'Resultado según cada modelo'));
        const grid = el('div', 'al-model-grid');
        P.modelos.slice().sort((a, b) => (a === MULTI ? -1 : b === MULTI ? 1 : 0)).forEach((m) => {
            const x = dato(g, m);
            const b = el('button', 'al-model' + (m === st.modelo ? ' is-sel' : ''));
            b.type = 'button';
            b.style.setProperty('--c', colorCat(x ? x.alerta : -1));
            b.append(el('span', 'al-model-name', etiquetaModelo(m)),
                el('span', 'al-model-cat', x && x.alerta >= 0 ? nombreCat(x.alerta) : 'Sin dato'),
                el('span', 'al-model-idx', x && x.alerta >= 0 ? `índice ${idx(x.indice)}` : (cubre(m) ? 'sin dato' : 'no cubre la ventana')));
            b.disabled = !(x && x.alerta >= 0);
            b.addEventListener('click', () => { st.modelo = m; $('sm-modelo').value = m; evaluarTodo(); });
            grid.append(b);
        });
        cont.append(grid);
    }

    // ---------------------------------------------------------------- exportar
    function slug(txt) {
        return txt.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'Nueva';
    }
    const nombreBase = () => `${slug(etiquetaCultivo(st.cultivo))}_${slug(st.nombre.trim() || 'Nueva_plaga')}`;

    function descargar(texto, nombre, tipo) {
        const blob = new Blob([texto], { type: tipo });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = nombre;
        document.body.append(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    }

    function aviso(txt) {
        const n = $('sm-aviso');
        n.textContent = txt;
        n.hidden = false;
        clearTimeout(aviso.t);
        aviso.t = setTimeout(() => { n.hidden = true; }, 4000);
    }

    async function copiarEnlace() {
        const url = enlace();
        try {
            await navigator.clipboard.writeText(url);
            aviso('Enlace copiado: quien lo abra verá esta misma configuración con el pronóstico del día.');
        } catch (e) {
            window.prompt('Copia este enlace:', url);
        }
    }

    function descargarCSV() {
        const act = activas();
        const q = (s) => `"${String(s).replace(/"/g, '""')}"`;
        const n = Math.min(st.ventana, P.fechas.length);
        const filas = [['plaga', 'cultivo', 'GID_3', 'departamento', 'municipio', 'ventana_inicio', 'ventana_fin', 'modelo',
            ...act.map(([k]) => VAR[k].v), ...act.map(([k]) => `cumple_${k}`), 'indice', 'alerta']];
        gids.forEach((g, i) => {
            P.modelos.forEach((m) => {
                const r = res[m][i], v = agg[m][i];
                filas.push([q(st.nombre), q(etiquetaCultivo(st.cultivo)), g, q(propsDe[g] ? propsDe[g].departamento : ''), q(nombreMun(g)),
                    P.fechas[0], P.fechas[n - 1], q(m),
                    ...act.map(([k]) => (nulo(v[VAR[k].v]) ? '' : +v[VAR[k].v].toFixed(2))),
                    ...act.map(([k]) => (r.conds[k] ?? '')),
                    nulo(r.indice) ? '' : +r.indice.toFixed(3), q(nombreCat(r.alerta))]);
            });
        });
        const cab = `# Simulación: ${act.map(([k, c]) => `${textoCond(k, c)} (peso ${c.peso})`).join('; ')}\n`;
        descargar('﻿' + cab + filas.map((r) => r.join(',')).join('\n'),
            `simulacion_${nombreBase().toLowerCase()}_${P.fecha_inicio}.csv`, 'text/csv;charset=utf-8');
    }

    function descargarJSON() {
        descargar(JSON.stringify(config(), null, 2), `simulacion_${nombreBase().toLowerCase()}.json`, 'application/json');
    }

    function abrirJSON(ev) {
        const f = ev.target.files[0];
        if (!f) return;
        f.text().then((t) => {
            cargarConfig(JSON.parse(t));
            sincronizarControles();
            recalcularTodo({ recrearCapa: true });
            aviso('Configuración cargada.');
        }).catch(() => aviso('No se pudo leer el archivo de configuración.'))
            .finally(() => { ev.target.value = ''; });
    }

    function descargarPython() {
        const act = activas();
        if (!act.length) { aviso('Activa al menos una variable antes de exportar.'); return; }
        const base = nombreBase();
        descargar(scriptPython(act, base), `Script_Indice_Riesgo_${base}_Pronostico.py`, 'text/x-python;charset=utf-8');
        aviso('Script descargado. Revísalo y agrégalo a Scripts/Modelos_Agroclimaticos_Python para incluirlo en las alertas diarias.');
    }

    function scriptPython(act, base) {
        const total = act.reduce((s, [, c]) => s + c.peso, 0);
        const plaga = st.nombre.trim() || 'Nueva plaga';
        const cultivo = etiquetaCultivo(st.cultivo);
        const id = base.toUpperCase();
        const py = (s) => JSON.stringify(s);
        const nombreCond = (k) => `cond_${k}`;
        const expr = (k, c) => {
            const v = VAR[k].v;
            if (c.op === 'entre') return `(${v} >= ${Math.min(c.a, c.b)}) & (${v} <= ${Math.max(c.a, c.b)})`;
            return `${v} ${OPS_PY[c.op]} ${c.a}`;
        };
        const umbral = act.map(([k, c]) => textoCond(k, c, true)).join(' y ');
        const entradas = act.map(([k]) => VAR[k].v);
        const ancho = Math.max(...act.map(([k]) => nombreCond(k).length));
        const hoy = new Date().toISOString().slice(0, 10);
        return `#!/usr/bin/env python3
## ============================================================================
## INDICE AGROCLIMATICO DE RIESGO - ${cultivo.toUpperCase()} / ${plaga.toUpperCase()}
## Sistema de Alerta Temprana de Plagas y Enfermedades - Bolivia
## Versión Python para el PRONÓSTICO a 16 días por municipio
## ============================================================================
##
## Generado con el Simulador de riesgo de la plataforma el ${hoy}.
## Es un modelo EXPLORATORIO: sus umbrales los definió el usuario y deben validarse
## con técnicos y literatura antes de usarlo como alerta oficial.
##
## Umbral de alerta:
##   ALERTA si: ${umbral}.
##
## Índice = suma de los pesos de las condiciones que se cumplen / ${total}
${act.map(([k, c]) => `##     (${textoCond(k, c, true)}) * ${c.peso}`).join('\n')}
##
## Semáforo: 0 = Normal; < 0.5 = Alerta baja; < 1 = Alerta media; 1 = Alerta alta.
##
## PARA INCLUIRLO EN LAS ALERTAS DIARIAS
##   Copiar este archivo en una carpeta nueva dentro de
##   Scripts/Modelos_Agroclimaticos_Python/ (p. ej. ${slug(cultivo)}/${slug(plaga)}/).
##   correr_modelos.py lo detecta solo y la web lo muestra en "Alertas".
##
## USO INDIVIDUAL
##   python ${`Script_Indice_Riesgo_${base}_Pronostico.py`} --pronostico data/pronostico/pronostico_latest.json
##
## Requiere solo numpy y pandas.
## ============================================================================

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================================
# 1. PARAMETROS
# ============================================================================

ID_MODELO     = ${py(id)}
CLAVE_CULTIVO = ${st.cultivo === TODOS ? 'None                 # None = todos los municipios del pronóstico' : py(st.cultivo)}
CULTIVO       = ${py(cultivo)}
PLAGA         = ${py(plaga)}

PERIODO_DIAS = ${st.ventana}
VENTANA_DIAS = PERIODO_DIAS             # ventana evaluada (primeros N días desde hoy)
UMBRAL_DIA_SECO_MM = 1                  # día seco: precipitación <= 1 mm
UMBRAL_DIA_LLUVIA_MM = 1                # día de lluvia: precipitación > 1 mm

# Raíz del repositorio: la primera carpeta superior que contiene data/geo
RAIZ = next((p for p in Path(__file__).resolve().parents if (p / "data" / "geo").is_dir()), Path.cwd())
RUTA_PRONOSTICO = RAIZ / "data" / "pronostico" / "pronostico_latest.json"
RUTA_CULTIVOS   = RAIZ / "data" / "geo" / "cultivos_municipios.json"
CARPETA_SALIDA  = RAIZ / "Salidas_Modelos_Agroclimaticos" / "Pronostico"

VAR_PRECIP, VAR_TMAX, VAR_TMIN, VAR_HR = (
    "precipitation_sum", "temperature_2m_max", "temperature_2m_min", "relative_humidity_2m")

CATEGORIAS = ["Normal", "Alerta baja", "Alerta media", "Alerta alta"]


# ============================================================================
# 2. FORMULA DEL INDICE
# ============================================================================

UMBRAL_TEXTO = ${py(umbral)}

# condición: (descripción, peso)
CONDICIONES = {
${act.map(([k, c]) => `    ${py(nombreCond(k))}: (${py(textoCond(k, c, true))}, ${c.peso}),`).join('\n')}
}


def calcular_indice(${entradas.join(', ')}):
    """Índice de riesgo ${cultivo} / ${plaga}. Devuelve (índice, condiciones)."""
${act.map(([k, c]) => `    ${nombreCond(k).padEnd(ancho)} = ${expr(k, c)}   # ${textoCond(k, c, true)}`).join('\n')}

    indice = (${act.map(([k, c]) => `${nombreCond(k)} * ${c.peso}`).join(' + ')}) / ${total}
    return indice, {${act.map(([k]) => `${py(nombreCond(k))}: ${nombreCond(k)}`).join(', ')}}


# Variables agregadas que necesita la fórmula (en el mismo orden de los argumentos)
ENTRADAS = ${py(entradas).replace(/,/g, ', ')}


# ============================================================================
# 3. FUNCIONES AUXILIARES (misma lógica que los modelos del SAT)
# ============================================================================

def calcular_racha_seca(dia_seco):
    """Días secos consecutivos, día a día. dia_seco: matriz municipios x días (1 = seco, 0 = lluvia)."""
    racha = np.zeros_like(dia_seco, dtype=float)
    actual = np.zeros(dia_seco.shape[0])
    seco = np.nan_to_num(dia_seco, nan=0.0)
    for d in range(dia_seco.shape[1]):
        actual = (actual + 1) * seco[:, d]
        racha[:, d] = actual
    return racha


def variables_del_periodo(tmax, tmin, precip, hr):
    """Agrega las series diarias de la ventana (municipios x días). Un día faltante deja el
    valor en NaN. Solo se calculan las variables que usa este modelo."""
    salida = {}
    if "t_media_periodo" in ENTRADAS:
        salida["t_media_periodo"] = ((tmax + tmin) / 2).mean(axis=1)
    if "tmax_media_periodo" in ENTRADAS:
        salida["tmax_media_periodo"] = tmax.mean(axis=1)
    if "tmin_media_periodo" in ENTRADAS:
        salida["tmin_media_periodo"] = tmin.mean(axis=1)
    if "precip_acum_periodo" in ENTRADAS:
        salida["precip_acum_periodo"] = precip.sum(axis=1)
    if "hr_media_periodo" in ENTRADAS:
        salida["hr_media_periodo"] = hr.mean(axis=1)
    if "dias_lluvia_periodo" in ENTRADAS:
        dia_lluvia = np.where(np.isnan(precip), np.nan, (precip > UMBRAL_DIA_LLUVIA_MM).astype(float))
        salida["dias_lluvia_periodo"] = dia_lluvia.sum(axis=1)
    if "dias_secos_periodo" in ENTRADAS:
        dia_seco = np.where(np.isnan(precip), np.nan, (precip <= UMBRAL_DIA_SECO_MM).astype(float))
        racha = calcular_racha_seca(dia_seco)[:, -1]           # racha vigente al último día
        racha[np.any(np.isnan(precip), axis=1)] = np.nan       # ventana incompleta -> sin dato
        salida["dias_secos_periodo"] = racha
    return salida


def clasificar_alerta(indice):
    """Semáforo. Devuelve 0..3 (índice en CATEGORIAS) o -1 si no hay dato."""
    i = np.round(indice, 9)
    cat = np.select([i == 0, i < 0.5, i < 1], [0, 1, 2], default=3)
    return np.where(np.isnan(indice), -1, cat).astype(int)


def _matriz(pronostico, variable, modelo, filas, n_dias):
    datos = pronostico["datos"][variable][modelo]
    m = [[np.nan if v is None else v for v in (datos[f][:n_dias] + [None] * n_dias)[:n_dias]] for f in filas]
    return np.array(m, dtype=float).reshape(len(filas), n_dias)


def calcular_alertas(pronostico, cultivos):
    """Calcula índice, semáforo, variables y condiciones para cada modelo del pronóstico."""
    pos = {g: i for i, g in enumerate(pronostico["municipios"])}
    if CLAVE_CULTIVO is None:
        municipios_cultivo = list(pronostico["municipios"])
    else:
        municipios_cultivo = list(cultivos[CLAVE_CULTIVO]["municipios"])
    gids = [g for g in municipios_cultivo if g in pos]
    sin_pronostico = [g for g in municipios_cultivo if g not in pos]
    filas = [pos[g] for g in gids]

    fechas = pronostico["fechas"]
    n = min(VENTANA_DIAS, len(fechas))

    resultado = {
        "cultivo": CULTIVO, "clave_cultivo": CLAVE_CULTIVO or "TODOS", "plaga": PLAGA,
        "periodo_dias_modelo": PERIODO_DIAS,
        "ventana": {"inicio": fechas[0], "fin": fechas[n - 1], "dias": n},
        "umbral": UMBRAL_TEXTO,
        "condiciones": {c: {"descripcion": d, "peso": p} for c, (d, p) in CONDICIONES.items()},
        "municipios": gids, "municipios_sin_pronostico": sin_pronostico,
        "modelos": {},
    }
    for modelo in pronostico["modelos"]:
        tmax = _matriz(pronostico, VAR_TMAX, modelo, filas, n)
        tmin = _matriz(pronostico, VAR_TMIN, modelo, filas, n)
        precip = _matriz(pronostico, VAR_PRECIP, modelo, filas, n)
        hr = _matriz(pronostico, VAR_HR, modelo, filas, n)

        variables = variables_del_periodo(tmax, tmin, precip, hr)
        with np.errstate(invalid="ignore"):
            indice, condiciones = calcular_indice(*[variables[v] for v in ENTRADAS])
        indice = np.asarray(indice, dtype=float)
        faltante = np.any([np.isnan(variables[v]) for v in ENTRADAS], axis=0)
        indice[faltante] = np.nan

        resultado["modelos"][modelo] = {
            "indice": indice,
            "alerta": clasificar_alerta(indice),
            "variables": variables,
            "condiciones": {c: np.where(faltante, np.nan, np.asarray(v, dtype=float))
                            for c, v in condiciones.items()},
        }
    return resultado


def a_tabla(resultado, cultivos):
    """Tabla larga: una fila por municipio y modelo."""
    nombres = cultivos[CLAVE_CULTIVO]["municipios"] if CLAVE_CULTIVO else {}
    filas = []
    for modelo, r in resultado["modelos"].items():
        for i, gid in enumerate(resultado["municipios"]):
            fila = {"GID_3": gid, "municipio": ", ".join(nombres.get(gid, {}).get("nombres", [])), "modelo": modelo,
                    "ventana_inicio": resultado["ventana"]["inicio"], "ventana_fin": resultado["ventana"]["fin"]}
            fila.update({v: r["variables"][v][i] for v in ENTRADAS})
            fila.update({c: r["condiciones"][c][i] for c in CONDICIONES})
            fila["indice"] = r["indice"][i]
            fila["alerta"] = CATEGORIAS[r["alerta"][i]] if r["alerta"][i] >= 0 else "Sin dato"
            filas.append(fila)
    return pd.DataFrame(filas)


# ============================================================================
# 4. EJECUCION INDIVIDUAL (genera un CSV)
# ============================================================================

def main():
    ap = argparse.ArgumentParser(description=f"Índice de riesgo {CULTIVO} / {PLAGA} sobre el pronóstico")
    ap.add_argument("--pronostico", type=Path, default=RUTA_PRONOSTICO)
    ap.add_argument("--cultivos", type=Path, default=RUTA_CULTIVOS)
    ap.add_argument("--salida", type=Path, default=CARPETA_SALIDA)
    args = ap.parse_args()

    with open(args.pronostico, encoding="utf-8") as f:
        pronostico = json.load(f)
    with open(args.cultivos, encoding="utf-8") as f:
        cultivos = json.load(f)

    resultado = calcular_alertas(pronostico, cultivos)
    tabla = a_tabla(resultado, cultivos)

    args.salida.mkdir(parents=True, exist_ok=True)
    ruta = args.salida / f"alerta_{ID_MODELO.lower()}_{pronostico['fecha_inicio']}.csv"
    tabla.round(3).to_csv(ruta, index=False, encoding="utf-8-sig")

    v = resultado["ventana"]
    print(f"{CULTIVO} / {PLAGA} | ventana {v['inicio']} a {v['fin']} ({v['dias']} días) | "
          f"{len(resultado['municipios'])} municipios")
    resumen = (tabla.groupby(["modelo", "alerta"]).size().unstack(fill_value=0)
               .reindex(columns=CATEGORIAS + ["Sin dato"], fill_value=0))
    print(resumen.to_string())
    print(f"Tabla guardada en: {ruta}")


if __name__ == "__main__":
    main()
`;
    }

    // ---------------------------------------------------------------- arranque
    document.addEventListener('DOMContentLoaded', () => {
        const panel = $('panel-simulador');
        if (!panel) return;
        const alMostrar = () => {
            if (panel.hidden) return;
            if (!iniciado) iniciar();
            else if (mapa) mapa.invalidateSize();
        };
        new MutationObserver(alMostrar).observe(panel, { attributes: true, attributeFilter: ['hidden'] });
        // Un enlace compartido (?sim=...) abre directamente el simulador
        if (new URLSearchParams(location.search).get('sim')) {
            setTimeout(() => {
                const b = document.getElementById('modo-simulador');
                if (b) b.click();
                const sec = document.getElementById('explorar');
                if (sec) sec.scrollIntoView();
            }, 0);
        }
        alMostrar();
    });
})();
