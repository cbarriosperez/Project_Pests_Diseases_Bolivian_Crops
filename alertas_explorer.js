/* ==========================================================================
   Explorador de Datos — Alertas de plagas y enfermedades (pronóstico)
   Datos: data/alertas/alertas_latest.json, generado a diario por
   Scripts/Modelos_Agroclimaticos_Python/correr_modelos.py a partir del
   pronóstico de 16 días + geometría en data/geo/.
   Requiere Leaflet 1.9 (cargado en index.html).
   ========================================================================== */
(function () {
    'use strict';

    const RUTAS = {
        alertas: 'data/alertas/alertas_latest.json',
        municipios: 'data/geo/municipios_bolivia.geojson',
        departamentos: 'data/geo/departamentos_bolivia.geojson',
        cultivos: 'data/geo/cultivos_municipios.json',
    };
    const MULTI = 'MULTIMODELO';
    const TZ = 'America/La_Paz';
    const COLOR_OSCURO = '#232d39';
    const SIN_DATO = '#e5e7eb';

    // Variables agregadas de la ventana y la condición que usa cada una
    const VARIABLES = {
        t_media_periodo:     { nombre: 'T media del periodo', unidad: '°C', dec: 1 },
        precip_acum_periodo: { nombre: 'Precipitación acumulada', unidad: 'mm', dec: 1 },
        hr_media_periodo:    { nombre: 'Humedad relativa media', unidad: '%', dec: 0 },
        dias_lluvia_periodo: { nombre: 'Días con lluvia (> 1 mm)', unidad: 'días', dec: 0 },
        dias_secos_periodo:  { nombre: 'Racha seca al final del periodo', unidad: 'días', dec: 0 },
    };
    const VAR_DE_CONDICION = {
        cond_t: 't_media_periodo', cond_precip: 'precip_acum_periodo', cond_hr: 'hr_media_periodo',
        cond_lluvia: 'dias_lluvia_periodo', cond_secos: 'dias_secos_periodo',
    };

    const st = { cultivo: null, ind: null, modelo: MULTI, gid: null };
    let A = null, MUN = null, DEP = null, CULT = null;
    let propsDe = {}, posDe = {};
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

    // ---------------------------------------------------------------- accesos a datos
    const IND = () => A.indicadores[st.ind];
    const etiquetaModelo = (m) => (m === MULTI ? 'Multimodelo' : m);
    const categoria = (c) => A.categorias[c];
    const colorCat = (c) => (c >= 0 ? categoria(c).color : SIN_DATO);
    const nombreCat = (c) => (c >= 0 ? categoria(c).nombre : 'Sin dato');
    const infoMun = (gid) => (CULT[IND().clave_cultivo] || { municipios: {} }).municipios[gid];
    const nombreMun = (gid) => {
        const i = infoMun(gid);
        return i ? i.nombres.join(' / ') : (propsDe[gid] ? propsDe[gid].municipio : gid);
    };
    function dato(gid, modelo = st.modelo) {
        const ind = IND(), i = posDe[st.ind][gid];
        if (i === undefined || !ind.alerta[modelo]) return null;
        return { alerta: ind.alerta[modelo][i], indice: ind.indice[modelo][i], i };
    }
    const modeloConDatos = (ind, m) => ind.alerta[m] && ind.alerta[m].some((c) => c >= 0);

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
            [A, MUN, DEP, CULT] = await Promise.all([
                cargarJSON(RUTAS.alertas, true), cargarJSON(RUTAS.municipios),
                cargarJSON(RUTAS.departamentos), cargarJSON(RUTAS.cultivos),
            ]);
        } catch (err) {
            iniciado = false;
            $('al-cargando').textContent =
                `No se pudieron cargar las alertas: ${err.message}. Recarga la página en unos minutos.`;
            return;
        }
        MUN.features.forEach((f) => { propsDe[f.properties.GID_3] = f.properties; });
        Object.entries(A.indicadores).forEach(([k, ind]) => {
            posDe[k] = {};
            ind.municipios.forEach((g, i) => { posDe[k][g] = i; });
        });

        $('al-cargando').hidden = true;
        $('al-contenido').hidden = false;
        poblarCultivos();
        mostrarCorrida();
        crearMapa();
        cambiarIndicador({ recrearCapa: true });
    }

    // ---------------------------------------------------------------- selectores
    function indicadoresDe(clave) {
        return Object.keys(A.indicadores).filter((k) => A.indicadores[k].clave_cultivo === clave);
    }

    function poblarCultivos() {
        const sc = $('al-cultivo');
        const claves = [...new Set(Object.values(A.indicadores).map((i) => i.clave_cultivo))];
        const etiqueta = (c) => (CULT[c] ? CULT[c].etiqueta : A.indicadores[indicadoresDe(c)[0]].cultivo);
        claves.sort((a, b) => etiqueta(a).localeCompare(etiqueta(b), 'es'))
            .forEach((c) => sc.add(new Option(etiqueta(c), c)));
        st.cultivo = claves.includes('PAPA') ? 'PAPA' : claves[0];
        sc.value = st.cultivo;
        poblarPlagas();

        sc.addEventListener('change', () => { st.cultivo = sc.value; poblarPlagas(); cambiarIndicador({ recrearCapa: true }); });
        $('al-plaga').addEventListener('change', () => { st.ind = $('al-plaga').value; cambiarIndicador({ recrearCapa: false }); });
        $('al-modelo').addEventListener('change', () => { st.modelo = $('al-modelo').value; actualizarTodo(); });
        $('al-csv').addEventListener('click', descargarCSV);
    }

    function poblarPlagas() {
        const sp = $('al-plaga');
        sp.innerHTML = '';
        const inds = indicadoresDe(st.cultivo);
        inds.forEach((k) => sp.add(new Option(A.indicadores[k].plaga, k)));
        st.ind = inds[0];
        sp.value = st.ind;
        sp.disabled = inds.length < 2;
    }

    function poblarModelos() {
        const sm = $('al-modelo'), ind = IND();
        sm.innerHTML = '';
        [MULTI, ...A.modelos.filter((m) => m !== MULTI)].forEach((m) => {
            const ok = modeloConDatos(ind, m);
            const o = new Option(etiquetaModelo(m) + (ok ? '' : ' (no cubre la ventana)'), m);
            o.disabled = !ok;
            sm.add(o);
        });
        if (!modeloConDatos(ind, st.modelo)) st.modelo = MULTI;
        sm.value = st.modelo;
    }

    function mostrarCorrida() {
        const fuente = A.pronostico_generado_utc || A.generado_utc;
        const hora = new Intl.DateTimeFormat('es-BO', {
            timeZone: TZ, day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
        }).format(new Date(fuente));
        $('al-corrida').innerHTML = '';
        $('al-corrida').append('Pronóstico del ', el('strong', '', hora), ' (hora de Bolivia).');

        const atraso = Math.round((aFecha(hoyBolivia()) - aFecha(A.fecha_inicio)) / 86400000);
        const alerta = $('al-alerta');
        if (atraso >= 2) {
            alerta.textContent = `Estas alertas usan el pronóstico del ${fmtLargo.format(aFecha(A.fecha_inicio))} ` +
                `y no se han actualizado en ${atraso} días.`;
            alerta.hidden = false;
        }
    }

    // ---------------------------------------------------------------- mapa
    function crearMapa() {
        mapa = L.map('al-mapa', { zoomSnap: 0.25, scrollWheelZoom: false, attributionControl: false });
        L.geoJSON(MUN, { interactive: false, style: { color: '#d5dbe2', weight: 0.5, fillColor: '#f1f3f5', fillOpacity: 1 } })
            .addTo(mapa);
        mapa.createPane('al-deptos').style.zIndex = 450;
        mapa.createPane('al-etiquetas').style.zIndex = 460;
        mapa.getPane('al-deptos').style.pointerEvents = 'none';
        mapa.getPane('al-etiquetas').style.pointerEvents = 'none';
        const deptos = L.geoJSON(DEP, { pane: 'al-deptos', interactive: false, style: { color: '#6b7785', weight: 1.2, fill: false } })
            .addTo(mapa);
        DEP.features.forEach((f) => {
            L.tooltip({ permanent: true, direction: 'center', className: 'pr-depto-label', pane: 'al-etiquetas' })
                .setLatLng(L.geoJSON(f).getBounds().getCenter()).setContent(f.properties.departamento).addTo(mapa);
        });
        mapa.fitBounds(deptos.getBounds(), { padding: [10, 10] });
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
        if (!d || d.alerta < 0) div.append('Sin dato para este modelo');
        else div.append(`${nombreCat(d.alerta)} (índice ${idx(d.indice)})`);
        return div;
    }

    function recrearCapa() {
        if (capa) mapa.removeLayer(capa);
        const gids = new Set(IND().municipios);
        capa = L.geoJSON(
            { type: 'FeatureCollection', features: MUN.features.filter((f) => gids.has(f.properties.GID_3)) },
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

    // ---------------------------------------------------------------- encabezado, resumen y leyenda
    function actualizarEncabezado() {
        const ind = IND();
        $('al-titulo').textContent = `Riesgo de ${ind.plaga.charAt(0).toLowerCase() + ind.plaga.slice(1)} en ${ind.cultivo.toLowerCase()}`;
        $('al-subtitulo').textContent = `${ind.municipios.length} municipios productores · ${etiquetaModelo(st.modelo).toLowerCase()}`;
        const v = ind.ventana;
        const f = $('al-ventana');
        f.innerHTML = '';
        f.append(`${corto(v.inicio)} – ${corto(v.fin)}`, el('small', '', `ventana de ${v.dias} días`));
    }

    function dibujarResumen() {
        const ind = IND(), cats = ind.alerta[st.modelo];
        const ha = (i) => { const m = infoMun(ind.municipios[i]); return m && m.sup_ha ? m.sup_ha : 0; };
        const totalHa = ind.municipios.reduce((s, _, i) => s + ha(i), 0);
        const cont = $('al-resumen');
        cont.innerHTML = '';
        [3, 2, 1, 0, -1].forEach((c) => {
            const idx = cats.map((x, i) => (x === c ? i : -1)).filter((i) => i >= 0);
            if (c === -1 && !idx.length) return;
            const sup = idx.reduce((s, i) => s + ha(i), 0);
            const t = el('div', 'al-tile' + (c < 0 ? ' al-tile-nd' : ''));
            t.style.setProperty('--c', colorCat(c));
            t.append(el('span', 'al-tile-n', String(idx.length)), el('span', 'al-tile-cat', nombreCat(c)));
            if (totalHa > 0 && c >= 0) t.append(el('span', 'al-tile-ha', `${num(100 * sup / totalHa, 0)} % de la superficie`));
            cont.append(t);
        });
    }

    function dibujarLeyenda() {
        const cont = $('al-leyenda');
        cont.innerHTML = '';
        cont.append(el('span', 'pr-legend-title', 'Nivel de alerta'));
        A.categorias.forEach((c) => {
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

    function dibujarFormula() {
        const ind = IND(), cont = $('al-formula');
        cont.innerHTML = '';
        cont.append(el('h4', '', 'Cómo se calcula'), el('p', 'al-umbral', `Alerta si: ${ind.umbral}.`));
        const ul = el('ul', 'al-cond-list');
        const total = Object.values(ind.condiciones).reduce((s, c) => s + c.peso, 0);
        Object.values(ind.condiciones).forEach((c) => {
            const li = el('li');
            li.append(el('span', '', c.descripcion), el('span', 'al-peso', `peso ${c.peso}`));
            ul.append(li);
        });
        cont.append(ul, el('p', 'al-nota',
            `Índice = suma de los pesos de las condiciones que se cumplen ÷ ${total}. ` +
            '0 = normal; menos de 0,5 = alerta baja; menos de 1 = alerta media; 1 = alerta alta.'));
        const a = el('a', '', 'Ver script del modelo');
        a.href = ind.script;
        a.target = '_blank';
        a.rel = 'noopener';
        cont.append(a);
    }

    // ---------------------------------------------------------------- ranking y detalle
    function dibujarRanking() {
        const ind = IND(), cont = $('al-ranking');
        cont.innerHTML = '';
        const filas = ind.municipios.map((g) => ({ g, d: dato(g), ha: (infoMun(g) || {}).sup_ha || 0 }))
            .filter((x) => x.d && x.d.alerta > 0)
            .sort((a, b) => b.d.indice - a.d.indice || b.ha - a.ha);
        if (!filas.length) {
            cont.append(el('p', 'al-vacio', 'Ningún municipio en alerta con este modelo.'));
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
        if (filas.length > 10) cont.append(el('p', 'al-nota', `y ${filas.length - 10} municipios más en alerta.`));
    }

    function dibujarDetalle() {
        const vacio = $('al-detalle-vacio'), cont = $('al-detalle');
        if (!st.gid || posDe[st.ind][st.gid] === undefined) {
            st.gid = null;
            vacio.hidden = false;
            cont.hidden = true;
            return;
        }
        vacio.hidden = true;
        cont.hidden = false;
        cont.innerHTML = '';
        const ind = IND(), g = st.gid, p = propsDe[g] || {}, info = infoMun(g) || {};
        const d = dato(g);

        const cab = el('div', 'al-det-head');
        const tit = el('div');
        tit.append(el('h4', '', nombreMun(g)));
        let meta = `${p.departamento || ''}${p.provincia ? ', provincia ' + p.provincia : ''}.`;
        if (info.sup_ha) meta += ` ${num(info.sup_ha, 0)} ha y ${num(info.prod_tn || 0, 0)} t de ${ind.cultivo.toLowerCase()} (censo agropecuario).`;
        tit.append(el('p', 'pr-series-meta', meta));
        const badge = el('span', 'al-badge', d && d.alerta >= 0 ? `${nombreCat(d.alerta)} · índice ${idx(d.indice)}` : 'Sin dato');
        badge.style.setProperty('--c', colorCat(d ? d.alerta : -1));
        cab.append(tit, badge);
        cont.append(cab);

        // Condiciones con el modelo elegido
        cont.append(el('h5', '', `Condiciones en la ventana · ${etiquetaModelo(st.modelo)}`));
        const tabla = el('table', 'al-cond-table');
        const thead = el('tr');
        ['Condición', 'Pronóstico', 'Peso', ''].forEach((h) => thead.append(el('th', '', h)));
        tabla.append(thead);
        Object.entries(ind.condiciones).forEach(([c, def]) => {
            const v = VAR_DE_CONDICION[c], meta = VARIABLES[v] || { unidad: '', dec: 1 };
            const valor = ind.variables[v] && ind.variables[v][st.modelo] ? ind.variables[v][st.modelo][d.i] : null;
            const cumple = ind.condiciones_cumplidas[c][st.modelo][d.i];
            const tr = el('tr');
            tr.append(el('td', '', def.descripcion),
                el('td', 'al-num', valor === null ? '—' : `${num(valor, meta.dec)} ${meta.unidad}`),
                el('td', 'al-num', String(def.peso)));
            const est = el('td', 'al-estado ' + (cumple === 1 ? 'si' : cumple === 0 ? 'no' : ''),
                cumple === 1 ? '✓ se cumple' : cumple === 0 ? '✗ no se cumple' : 'sin dato');
            tr.append(est);
            tabla.append(tr);
        });
        cont.append(tabla);

        // Alerta en cada modelo
        cont.append(el('h5', '', 'Alerta según cada modelo'));
        const grid = el('div', 'al-model-grid');
        [MULTI, ...A.modelos.filter((m) => m !== MULTI)].forEach((m) => {
            const x = dato(g, m);
            const b = el('button', 'al-model' + (m === st.modelo ? ' is-sel' : ''));
            b.type = 'button';
            b.style.setProperty('--c', colorCat(x ? x.alerta : -1));
            b.append(el('span', 'al-model-name', etiquetaModelo(m)),
                el('span', 'al-model-cat', x && x.alerta >= 0 ? nombreCat(x.alerta) : 'Sin dato'),
                el('span', 'al-model-idx', x && x.alerta >= 0 ? `índice ${idx(x.indice)}` : 'no cubre la ventana'));
            b.disabled = !(x && x.alerta >= 0);
            b.title = b.disabled ? '' : `Ver el mapa con ${etiquetaModelo(m)}`;
            b.addEventListener('click', () => { st.modelo = m; $('al-modelo').value = m; actualizarTodo(); });
            grid.append(b);
        });
        cont.append(grid);
        const conDato = A.modelos.filter((m) => m !== MULTI).map((m) => dato(g, m)).filter((x) => x && x.alerta >= 0);
        const enAlerta = conDato.filter((x) => x.alerta >= 2).length;
        if (conDato.length) {
            cont.append(el('p', 'al-nota',
                `${enAlerta} de ${conDato.length} modelos pronostican alerta media o alta en este municipio.`));
        }
    }

    // ---------------------------------------------------------------- CSV
    function descargarCSV() {
        const ind = IND();
        const vars = Object.keys(ind.variables), conds = Object.keys(ind.condiciones);
        const q = (s) => `"${String(s).replace(/"/g, '""')}"`;
        const filas = [['cultivo', 'plaga', 'GID_3', 'departamento', 'municipio', 'ventana_inicio', 'ventana_fin', 'modelo',
            ...vars, ...conds, 'indice', 'alerta']];
        ind.municipios.forEach((g, i) => {
            A.modelos.forEach((m) => {
                if (!ind.alerta[m]) return;
                const c = ind.alerta[m][i];
                filas.push([q(ind.cultivo), q(ind.plaga), g, q(propsDe[g] ? propsDe[g].departamento : ''), q(nombreMun(g)),
                    ind.ventana.inicio, ind.ventana.fin, q(m),
                    ...vars.map((v) => (ind.variables[v][m][i] ?? '')),
                    ...conds.map((k) => (ind.condiciones_cumplidas[k][m][i] ?? '')),
                    ind.indice[m][i] ?? '', q(nombreCat(c))]);
            });
        });
        const blob = new Blob(['﻿' + filas.map((r) => r.join(',')).join('\n')], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `alertas_${st.ind.toLowerCase()}_${A.fecha_inicio}.csv`;
        document.body.append(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    }

    // ---------------------------------------------------------------- orquestación
    function cambiarIndicador({ recrearCapa: recrear }) {
        poblarModelos();
        if (st.gid && posDe[st.ind][st.gid] === undefined) st.gid = null;
        if (recrear) recrearCapa();
        dibujarFormula();
        actualizarTodo();
    }

    function actualizarTodo() {
        capa.setStyle(estilo);
        actualizarEncabezado();
        dibujarResumen();
        dibujarLeyenda();
        dibujarRanking();
        dibujarDetalle();
        const ind = IND();
        const sin = A.modelos.filter((m) => m !== MULTI && !modeloConDatos(ind, m));
        $('al-nota-modelos').textContent = sin.length
            ? `${sin.join(', ')} no ${sin.length > 1 ? 'cubren' : 'cubre'} los ${ind.ventana.dias} días de la ventana, por eso no ${sin.length > 1 ? 'tienen' : 'tiene'} índice.`
            : '';
    }

    // Se inicia la primera vez que el panel se hace visible (el cambio de pestaña lo maneja pronostico_explorer.js)
    document.addEventListener('DOMContentLoaded', () => {
        const panel = $('panel-alertas');
        if (!panel) return;
        const alMostrar = () => {
            if (panel.hidden) return;
            if (!iniciado) iniciar();
            else if (mapa) mapa.invalidateSize();
        };
        new MutationObserver(alMostrar).observe(panel, { attributes: true, attributeFilter: ['hidden'] });
        alMostrar();
    });
})();
