/* ==========================================================================
   Explorador de Datos — Pronóstico a 16 días por municipio productor
   Datos: data/pronostico/pronostico_latest.json (actualizado a diario por
   .github/workflows/pronostico_diario.yml) + geometría en data/geo/.
   Requiere Leaflet 1.9 y Chart.js 4 (cargados en index.html).
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
    const TZ = 'America/La_Paz';
    const COLOR_OSCURO = '#232d39';

    // Escalas fijas por variable: el mismo color significa lo mismo en los 16 días.
    // "cortes" son los límites entre clases; hay cortes.length + 1 colores.
    const ESCALAS = {
        precipitation_sum: {
            cortes: [1, 2, 5, 10, 20, 30],
            colores: ['#f7fbff', '#deebf7', '#c6dbef', '#9ecae1', '#6baed6', '#3182bd', '#08519c'],
        },
        temperature_2m_max: {
            cortes: [5, 10, 15, 20, 25, 30, 35],
            colores: ['#ffffcc', '#ffeda0', '#fed976', '#feb24c', '#fd8d3c', '#fc4e2a', '#e31a1c', '#b10026'],
        },
        temperature_2m_min: {
            cortes: [-5, 0, 5, 10, 15, 20, 25],
            colores: ['#2166ac', '#4393c3', '#92c5de', '#d1e5f0', '#fddbc7', '#f4a582', '#d6604d', '#b2182b'],
        },
        relative_humidity_2m: {
            cortes: [20, 30, 40, 50, 60, 70, 80, 90],
            colores: ['#8c510a', '#bf812d', '#dfc27d', '#f6e8c3', '#f0f0f0', '#c7eae5', '#80cdc1', '#35978f', '#01665e'],
        },
    };

    const st = { cultivo: null, variable: 'precipitation_sum', modelo: MULTI, dia: 0, gid: null, timer: null };
    let D = null, MUN = null, DEP = null, CULT = null;
    let filaDe = {}, propsDe = {};
    let mapa = null, capaCultivo = null, grafico = null;
    let iniciado = false;

    const $ = (id) => document.getElementById(id);
    const num = (v, dec = 1) => v.toLocaleString('es-BO', { maximumFractionDigits: dec, minimumFractionDigits: 0 });

    // ---------------------------------------------------------------- fechas
    const fmtDiaSemana = new Intl.DateTimeFormat('es-BO', { weekday: 'short', timeZone: 'UTC' });
    const fmtLargo = new Intl.DateTimeFormat('es-BO', { weekday: 'long', day: 'numeric', month: 'long', timeZone: 'UTC' });
    const fmtCorto = new Intl.DateTimeFormat('es-BO', { day: 'numeric', month: 'short', timeZone: 'UTC' });
    const aFecha = (iso) => new Date(iso + 'T12:00:00Z');
    const hoyBolivia = () => new Intl.DateTimeFormat('en-CA', { timeZone: TZ }).format(new Date());

    // ---------------------------------------------------------------- utilidades de datos
    const etiquetaModelo = (m) => (m === MULTI ? 'Multimodelo (promedio)' : m);
    const unidad = () => D.variables[st.variable].unidad;
    const valor = (gid, modelo = st.modelo, dia = st.dia, variable = st.variable) => {
        const f = filaDe[gid];
        if (f === undefined) return null;
        const v = D.datos[variable][modelo][f][dia];
        return v === undefined ? null : v;
    };
    const colorDe = (v) => {
        const e = ESCALAS[st.variable];
        let i = 0;
        while (i < e.cortes.length && v >= e.cortes[i]) i++;
        return e.colores[i];
    };
    const municipiosCultivo = () => CULT[st.cultivo].municipios;

    // Para cada día: ¿hay dato para la variable/modelo actuales? y, en el multimodelo, cuántos modelos aportan
    function disponibilidad() {
        const n = D.fechas.length;
        if (st.modelo === MULTI) {
            const cuenta = D.n_modelos_multimodelo[st.variable];
            return Array.from({ length: n }, (_, d) => ({ ok: cuenta[d] > 0, modelos: cuenta[d] }));
        }
        const mat = D.datos[st.variable][st.modelo];
        return Array.from({ length: n }, (_, d) => ({ ok: mat.some((fila) => fila[d] !== null), modelos: null }));
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
        const cont = $('pr-contenido');
        try {
            if (typeof L === 'undefined') throw new Error('no se cargó la librería de mapas (Leaflet)');
            [D, MUN, DEP, CULT] = await Promise.all([
                cargarJSON(RUTAS.pronostico, true), cargarJSON(RUTAS.municipios),
                cargarJSON(RUTAS.departamentos), cargarJSON(RUTAS.cultivos),
            ]);
        } catch (err) {
            iniciado = false;
            $('pr-cargando').textContent =
                `No se pudo cargar el pronóstico: ${err.message}. Recarga la página en unos minutos.`;
            return;
        }
        D.municipios.forEach((g, i) => { filaDe[g] = i; });
        MUN.features.forEach((f) => { propsDe[f.properties.GID_3] = f.properties; });

        $('pr-cargando').hidden = true;
        cont.hidden = false;

        poblarSelectores();
        mostrarCorrida();
        crearMapa();
        crearTiraDias();
        actualizarTodo({ recrearCapa: true });
    }

    // ---------------------------------------------------------------- selectores
    function poblarSelectores() {
        const sc = $('pr-cultivo'), sv = $('pr-variable'), sm = $('pr-modelo');
        const cultivos = Object.keys(CULT).sort((a, b) => CULT[a].etiqueta.localeCompare(CULT[b].etiqueta, 'es'));
        cultivos.forEach((c) => sc.add(new Option(CULT[c].etiqueta, c)));
        st.cultivo = cultivos.includes('PAPA') ? 'PAPA' : cultivos[0];
        sc.value = st.cultivo;

        Object.keys(D.variables).forEach((v) =>
            sv.add(new Option(`${D.variables[v].nombre} (${D.variables[v].unidad})`, v)));
        sv.value = st.variable;

        [MULTI, ...D.modelos.filter((m) => m !== MULTI)].forEach((m) => sm.add(new Option(etiquetaModelo(m), m)));
        sm.value = st.modelo;

        sc.addEventListener('change', () => {
            st.cultivo = sc.value;
            if (st.gid && !municipiosCultivo()[st.gid]) st.gid = null;
            actualizarTodo({ recrearCapa: true });
        });
        sv.addEventListener('change', () => { st.variable = sv.value; actualizarTodo(); });
        sm.addEventListener('change', () => { st.modelo = sm.value; actualizarTodo(); });
        $('pr-csv').addEventListener('click', descargarCSV);
    }

    function mostrarCorrida() {
        const generado = new Date(D.generado_utc);
        const hora = new Intl.DateTimeFormat('es-BO', {
            timeZone: TZ, day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
        }).format(generado);
        const modelos = D.modelos.filter((m) => m !== MULTI).join(', ');
        let html = `Corrida del <strong>${hora}</strong> (hora de Bolivia).<br>Modelos: ${modelos}.`;
        const fallidos = Object.keys(D.modelos_fallidos || {});
        if (fallidos.length) html += `<br>Sin datos hoy: ${fallidos.join(', ')}.`;
        $('pr-corrida').innerHTML = html;

        const atraso = Math.round((aFecha(hoyBolivia()) - aFecha(D.fecha_inicio)) / 86400000);
        const alerta = $('pr-alerta');
        if (atraso >= 2) {
            alerta.textContent = `Este pronóstico empieza el ${fmtLargo.format(aFecha(D.fecha_inicio))} y no se ha ` +
                `actualizado en ${atraso} días. Los primeros días ya pasaron.`;
            alerta.hidden = false;
        }
    }

    // ---------------------------------------------------------------- mapa
    function crearMapa() {
        mapa = L.map('pr-mapa', { zoomSnap: 0.25, scrollWheelZoom: false, attributionControl: false });

        L.geoJSON(MUN, {
            interactive: false,
            style: { color: '#d5dbe2', weight: 0.5, fillColor: '#f1f3f5', fillOpacity: 1 },
        }).addTo(mapa);

        const deptos = L.geoJSON(DEP, {
            interactive: false,
            style: { color: '#7b8794', weight: 1.2, fill: false },
        });
        mapa.createPane('deptos').style.zIndex = 450;
        mapa.createPane('etiquetas').style.zIndex = 460;
        mapa.getPane('deptos').style.pointerEvents = 'none';
        mapa.getPane('etiquetas').style.pointerEvents = 'none';
        L.geoJSON(DEP, { pane: 'deptos', interactive: false, style: { color: '#6b7785', weight: 1.2, fill: false } })
            .addTo(mapa);
        DEP.features.forEach((f) => {
            const centro = L.geoJSON(f).getBounds().getCenter();
            L.tooltip({ permanent: true, direction: 'center', className: 'pr-depto-label', pane: 'etiquetas' })
                .setLatLng(centro).setContent(f.properties.departamento).addTo(mapa);
        });

        mapa.fitBounds(deptos.getBounds(), { padding: [10, 10] });
    }

    function estilo(feature) {
        const gid = feature.properties.GID_3;
        const v = valor(gid);
        const seleccionado = gid === st.gid;
        const base = {
            color: seleccionado ? COLOR_OSCURO : '#6b7785',
            weight: seleccionado ? 3 : 0.6,
            fillOpacity: 1,
        };
        if (v === null) return { ...base, fillColor: '#e5e7eb', dashArray: '3' };
        return { ...base, fillColor: colorDe(v), dashArray: null };
    }

    function contenidoTooltip(capa) {
        const gid = capa.feature.properties.GID_3;
        const p = propsDe[gid];
        const nombres = municipiosCultivo()[gid].nombres.join(' / ');
        const v = valor(gid);
        const texto = v === null ? 'sin dato' : `${num(v)} ${unidad()}`;
        const div = document.createElement('div');
        div.innerHTML = '<b></b><br><span></span><br>';
        div.querySelector('b').textContent = nombres;
        div.querySelector('span').textContent = p.departamento;
        div.append(`${D.variables[st.variable].nombre}: ${texto}`);
        return div;
    }

    function recrearCapaCultivo() {
        if (capaCultivo) mapa.removeLayer(capaCultivo);
        const gids = municipiosCultivo();
        capaCultivo = L.geoJSON(
            { type: 'FeatureCollection', features: MUN.features.filter((f) => gids[f.properties.GID_3]) },
            {
                style: estilo,
                onEachFeature: (f, capa) => {
                    capa.bindTooltip(() => contenidoTooltip(capa), { sticky: true, className: 'pr-tooltip' });
                    capa.on('click', () => seleccionarMunicipio(f.properties.GID_3));
                },
            },
        ).addTo(mapa);
    }

    function seleccionarMunicipio(gid) {
        st.gid = gid;
        capaCultivo.setStyle(estilo);
        capaCultivo.eachLayer((c) => { if (c.feature.properties.GID_3 === gid) c.bringToFront(); });
        dibujarSerie();
    }

    // ---------------------------------------------------------------- leyenda
    function dibujarLeyenda() {
        const e = ESCALAS[st.variable];
        const cont = $('pr-leyenda');
        cont.innerHTML = '';
        const titulo = document.createElement('span');
        titulo.className = 'pr-legend-title';
        titulo.textContent = `${D.variables[st.variable].nombre} (${unidad()})`;
        cont.append(titulo);
        e.colores.forEach((c, i) => {
            const desde = i === 0 ? null : e.cortes[i - 1];
            const hasta = i < e.cortes.length ? e.cortes[i] : null;
            const etiqueta = desde === null ? `< ${hasta}` : hasta === null ? `≥ ${desde}` : `${desde}–${hasta}`;
            const item = document.createElement('div');
            item.className = 'pr-legend-item';
            item.innerHTML = `<span class="pr-legend-swatch" style="background:${c}"></span><span></span>`;
            item.lastChild.textContent = etiqueta;
            cont.append(item);
        });
        const nd = document.createElement('div');
        nd.className = 'pr-legend-item pr-legend-nodata';
        nd.innerHTML = '<span class="pr-legend-swatch"></span><span>sin dato</span>';
        cont.append(nd);
    }

    // ---------------------------------------------------------------- tira de días
    function crearTiraDias() {
        const tira = $('pr-dias');
        D.fechas.forEach((iso, d) => {
            const b = document.createElement('button');
            b.type = 'button';
            b.className = 'pr-day';
            b.dataset.dia = d;
            const f = aFecha(iso);
            b.innerHTML = `<span class="wd"></span><span class="dn"></span><span class="conf"></span>`;
            b.querySelector('.wd').textContent = fmtDiaSemana.format(f).replace('.', '').slice(0, 2);
            b.querySelector('.dn').textContent = f.getUTCDate();
            b.addEventListener('click', () => { detener(); irADia(d); });
            tira.append(b);
        });
        tira.addEventListener('keydown', (ev) => {
            if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
            ev.preventDefault();
            detener();
            const disp = disponibilidad();
            const paso = ev.key === 'ArrowRight' ? 1 : -1;
            for (let d = st.dia + paso; d >= 0 && d < disp.length; d += paso) {
                if (disp[d].ok) { irADia(d); tira.children[d].focus(); break; }
            }
        });
        $('pr-play').addEventListener('click', () => (st.timer ? detener() : reproducir()));
    }

    function actualizarTiraDias() {
        const disp = disponibilidad();
        const total = D.modelos.length - 1;
        [...$('pr-dias').children].forEach((b, d) => {
            const info = disp[d];
            b.disabled = !info.ok;
            b.setAttribute('aria-pressed', String(d === st.dia));
            b.classList.toggle('is-weak', st.modelo === MULTI && info.ok && info.modelos < 3);
            const conf = b.querySelector('.conf');
            conf.innerHTML = '';
            if (st.modelo === MULTI) {
                for (let i = 0; i < total; i++) {
                    const p = document.createElement('i');
                    if (i >= info.modelos) p.className = 'off';
                    conf.append(p);
                }
            }
            const fecha = fmtLargo.format(aFecha(D.fechas[d]));
            b.title = !info.ok ? `${fecha}: sin dato para este modelo`
                : st.modelo === MULTI ? `${fecha}: promedio de ${info.modelos} de ${total} modelos` : fecha;
            b.setAttribute('aria-label', b.title);
        });

        const nota = $('pr-dias-nota');
        if (st.modelo === MULTI) {
            const cuenta = D.n_modelos_multimodelo[st.variable];
            const debil = cuenta.findIndex((n) => n < 3);
            nota.textContent = 'Los puntos indican cuántos modelos entran en el promedio de cada día.' +
                (debil >= 0 ? ` Desde el día ${debil + 1} el promedio usa menos de 3 modelos.` : '');
        } else {
            const dias = disp.filter((x) => x.ok).length;
            nota.textContent = `${st.modelo} cubre ${dias} de los ${D.fechas.length} días del horizonte.`;
        }
    }

    function irADia(d) {
        st.dia = d;
        capaCultivo.setStyle(estilo);
        actualizarTiraDias();
        actualizarEncabezado();
        if (grafico) grafico.update('none');
    }

    function reproducir() {
        const disp = disponibilidad();
        const habiles = disp.map((x, d) => (x.ok ? d : -1)).filter((d) => d >= 0);
        if (!habiles.length) return;
        if (st.dia >= habiles[habiles.length - 1]) irADia(habiles[0]);
        $('pr-play').textContent = '❚❚';
        $('pr-play').setAttribute('aria-label', 'Pausar recorrido por días');
        st.timer = setInterval(() => {
            const sig = habiles.find((d) => d > st.dia);
            if (sig === undefined) { detener(); return; }
            irADia(sig);
        }, 900);
    }

    function detener() {
        if (!st.timer) return;
        clearInterval(st.timer);
        st.timer = null;
        $('pr-play').textContent = '▶';
        $('pr-play').setAttribute('aria-label', 'Recorrer los 16 días');
    }

    // ---------------------------------------------------------------- encabezado
    function actualizarEncabezado() {
        const n = Object.keys(municipiosCultivo()).length;
        $('pr-titulo').textContent = `${D.variables[st.variable].nombre}, ${etiquetaModelo(st.modelo).toLowerCase()}`;
        $('pr-subtitulo').textContent =
            `${n} municipios productores de ${CULT[st.cultivo].etiqueta.toLowerCase()}`;
        const f = aFecha(D.fechas[st.dia]);
        const fecha = fmtLargo.format(f);
        $('pr-fecha').innerHTML = '';
        $('pr-fecha').append(fecha.charAt(0).toUpperCase() + fecha.slice(1));
        const small = document.createElement('small');
        small.textContent = `día ${st.dia + 1} de ${D.fechas.length}`;
        $('pr-fecha').append(small);
    }

    // ---------------------------------------------------------------- serie del municipio
    const lineaDia = {
        id: 'lineaDia',
        afterDatasetsDraw(chart) {
            const x = chart.scales.x.getPixelForValue(st.dia);
            const { top, bottom } = chart.chartArea;
            const ctx = chart.ctx;
            ctx.save();
            ctx.strokeStyle = 'rgba(196, 22, 27, 0.7)';
            ctx.setLineDash([4, 3]);
            ctx.lineWidth = 1.5;
            ctx.beginPath();
            ctx.moveTo(x, top);
            ctx.lineTo(x, bottom);
            ctx.stroke();
            ctx.restore();
        },
    };

    function dibujarSerie() {
        const vacio = $('pr-serie-vacia'), lleno = $('pr-serie-contenido');
        if (!st.gid) {
            vacio.hidden = false;
            lleno.hidden = true;
            if (grafico) { grafico.destroy(); grafico = null; }
            return;
        }
        vacio.hidden = true;
        lleno.hidden = false;

        const p = propsDe[st.gid];
        const info = municipiosCultivo()[st.gid];
        $('pr-serie-titulo').textContent = info.nombres.join(' / ');
        $('pr-serie-meta').textContent =
            `${p.departamento}, provincia ${p.provincia}. ${num(info.sup_ha, 0)} ha y ` +
            `${num(info.prod_tn, 0)} t de ${CULT[st.cultivo].etiqueta.toLowerCase()} (censo agropecuario).`;

        if (typeof Chart === 'undefined') {
            $('pr-grafico').replaceWith(Object.assign(document.createElement('p'),
                { textContent: 'No se pudo cargar la librería de gráficos.' }));
            return;
        }

        const esPrecip = st.variable === 'precipitation_sum';
        const etiquetas = D.fechas.map((iso) => fmtCorto.format(aFecha(iso)).replace('.', ''));
        const modelos = [...D.modelos.filter((m) => m !== MULTI), MULTI];
        const datasets = modelos.map((m) => {
            const datos = D.fechas.map((_, d) => valor(st.gid, m, d));
            const elegido = m === st.modelo;
            if (elegido && esPrecip) {
                return { type: 'bar', label: etiquetaModelo(m), data: datos, backgroundColor: '#3182bd', order: 2 };
            }
            if (elegido) {
                return { label: etiquetaModelo(m), data: datos, borderColor: COLOR_OSCURO, backgroundColor: COLOR_OSCURO,
                    borderWidth: 3, pointRadius: 2.5, order: 0, spanGaps: false };
            }
            if (m === MULTI) {
                return { label: etiquetaModelo(m), data: datos, borderColor: '#4a8031', borderDash: [6, 4],
                    borderWidth: 2, pointRadius: 0, order: 1 };
            }
            return { label: m, data: datos, borderColor: '#b6c0cc', borderWidth: 1.2, pointRadius: 0, order: 3 };
        });

        if (grafico) grafico.destroy();
        grafico = new Chart($('pr-grafico'), {
            type: 'line',
            data: { labels: etiquetas, datasets },
            plugins: [lineaDia],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { position: 'bottom', labels: { boxWidth: 14, boxHeight: 2, font: { size: 11 } } },
                    tooltip: {
                        callbacks: {
                            label: (c) => `${c.dataset.label}: ${c.parsed.y === null ? 'sin dato' : num(c.parsed.y) + ' ' + unidad()}`,
                        },
                    },
                },
                scales: {
                    x: { grid: { display: false }, ticks: { font: { size: 11 }, maxRotation: 0, autoSkipPadding: 6 } },
                    y: {
                        beginAtZero: esPrecip,
                        title: { display: true, text: `${D.variables[st.variable].nombre} (${unidad()})`, font: { size: 11 } },
                        ticks: { font: { size: 11 } },
                    },
                },
                onClick: (_, elementos) => {
                    if (elementos.length) { detener(); irADia(elementos[0].index); }
                },
            },
        });
    }

    // ---------------------------------------------------------------- CSV
    function descargarCSV() {
        const gids = Object.keys(municipiosCultivo());
        const filas = [['cultivo', 'GID_3', 'departamento', 'municipio', 'fecha', 'variable', 'unidad', 'modelo', 'valor']];
        const q = (s) => `"${String(s).replace(/"/g, '""')}"`;
        gids.forEach((gid) => {
            const p = propsDe[gid];
            const nombre = municipiosCultivo()[gid].nombres.join(' / ');
            Object.keys(D.variables).forEach((v) => {
                D.modelos.forEach((m) => {
                    D.fechas.forEach((f, d) => {
                        const x = valor(gid, m, d, v);
                        filas.push([q(CULT[st.cultivo].etiqueta), gid, q(p.departamento), q(nombre), f, v,
                            q(D.variables[v].unidad), q(m), x === null ? '' : x]);
                    });
                });
            });
        });
        const blob = new Blob(['\ufeff' + filas.map((r) => r.join(',')).join('\n')], { type: 'text/csv;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `pronostico_${st.cultivo.toLowerCase()}_${D.fecha_inicio}.csv`;
        document.body.append(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    }

    // ---------------------------------------------------------------- orquestación
    function actualizarTodo({ recrearCapa = false } = {}) {
        const disp = disponibilidad();
        if (!disp[st.dia].ok) {
            const previo = disp.map((x, d) => (x.ok && d <= st.dia ? d : -1)).filter((d) => d >= 0).pop();
            st.dia = previo !== undefined ? previo : Math.max(0, disp.findIndex((x) => x.ok));
        }
        if (recrearCapa) recrearCapaCultivo(); else capaCultivo.setStyle(estilo);
        dibujarLeyenda();
        actualizarTiraDias();
        actualizarEncabezado();
        dibujarSerie();
    }

    // ---------------------------------------------------------------- selector de modo
    function mostrarModo(modo) {
        document.querySelectorAll('.explorer-mode-btn').forEach((b) => {
            const activo = b.dataset.modo === modo;
            b.setAttribute('aria-selected', String(activo));
            b.tabIndex = activo ? 0 : -1;
        });
        document.querySelectorAll('.explorer-panel').forEach((p) => { p.hidden = p.dataset.modo !== modo; });
        if (modo === 'pronostico') {
            if (!iniciado) iniciar();
            else if (mapa) mapa.invalidateSize();
        } else {
            detener();
        }
    }

    document.addEventListener('DOMContentLoaded', () => {
        const botones = [...document.querySelectorAll('.explorer-mode-btn')];
        if (!botones.length) return;
        botones.forEach((b, i) => {
            b.addEventListener('click', () => mostrarModo(b.dataset.modo));
            b.addEventListener('keydown', (ev) => {
                if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
                const otro = botones[(i + (ev.key === 'ArrowRight' ? 1 : -1) + botones.length) % botones.length];
                otro.focus();
                mostrarModo(otro.dataset.modo);
            });
        });
        const inicial = botones.find((b) => b.getAttribute('aria-selected') === 'true') || botones[0];
        mostrarModo(inicial.dataset.modo);
    });
})();
