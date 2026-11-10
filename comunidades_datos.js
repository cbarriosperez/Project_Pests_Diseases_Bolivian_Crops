/* ==========================================================================
   Explorador de Datos — Nivel «Comunidad»
   Módulo compartido por los exploradores de pronóstico, alertas y simulador.

   El pronóstico por comunidades se descarga a diario para una grilla de celdas
   (data/pronostico/pronostico_comunidades_latest.json, 0,25° con la API gratuita
   de Open-Meteo). Cada comunidad toma el pronóstico de su celda y su temperatura
   máxima y mínima se corrigen por la diferencia de altitud entre el punto de la
   celda y la comunidad (6,5 °C por km):
       T_comunidad = T_celda + 0,0065 × (z_celda − z_comunidad)
   La precipitación y la humedad relativa se toman tal cual de la celda.

   Las alertas por comunidad se calculan aquí, en el navegador, con el mismo
   núcleo del simulador (window.SimuladorRiesgo) y los umbrales de los 15
   indicadores; el resultado es idéntico al de
   Scripts/Modelos_Agroclimaticos_Python/correr_modelos_comunidades.py.

   Expone window.SATComunidades:
     nivel                      'municipios' | 'comunidades'
     cambiarNivel(n)            cambia el nivel y avisa con el evento 'sat:nivel'
     cargar(MUN, CULT)          promesa con { P, MUN, CULT, propsDe, filaDe, info }
     alertas(Amun)              objeto con el formato de alertas_latest.json, por comunidad
   ========================================================================== */
(function () {
    'use strict';

    const RUTAS = {
        pronostico: 'data/pronostico/pronostico_comunidades_latest.json',
        comunidades: 'data/geo/comunidades/comunidades_bolivia.geojson',
        elevacion: 'data/geo/comunidades/elevacion_comunidades.json',
    };
    const MULTI = 'MULTIMODELO';
    const VARIABLES_TEMPERATURA = ['temperature_2m_max', 'temperature_2m_min'];

    let nivel = 'municipios';
    let promesa = null;      // carga en curso o terminada
    let datos = null;
    const cacheAlertas = new WeakMap();

    async function cargarJSON(url, sinCache) {
        const r = await fetch(url, sinCache ? { cache: 'no-cache' } : {});
        if (!r.ok) throw new Error(r.status === 404 ? 'todavía no se ha generado el pronóstico por comunidades' : `${url} respondió ${r.status}`);
        return r.json();
    }

    // Misma regla que clave_celda() en pronostico_municipios.py
    const claveCelda = (lat, lon, res) => (res > 0 ? `${Math.floor(lat / res)}_${Math.floor(lon / res)}` : null);

    async function construir(MUNmun, CULTmun) {
        const [PC, GEO, ELEV] = await Promise.all([
            cargarJSON(RUTAS.pronostico, true),
            cargarJSON(RUTAS.comunidades),
            cargarJSON(RUTAS.elevacion).catch(() => null),   // opcional: sin él no se corrige la temperatura
        ]);
        const z = (ELEV && ELEV.z) || {};
        const res = PC.resolucion_grados;
        const gradiente = (PC.gradiente_temp_c_por_km ?? 6.5) / 1000;
        const filaPunto = {};
        PC.municipios.forEach((k, i) => { filaPunto[k] = i; });
        const zPunto = PC.puntos_elevacion_m || [];
        const propsMun = {};
        MUNmun.features.forEach((f) => { propsMun[f.properties.GID_3] = f.properties; });

        const ids = [], filas = [], ajustes = [], features = [], propsDe = {}, filaDe = {};
        GEO.features.forEach((f) => {
            const p = f.properties;
            const clave = claveCelda(p.y, p.x, res) ?? p.id;
            const i = filaPunto[clave];
            if (i === undefined) return;
            const zc = z[p.id], zp = zPunto[i];
            const corregida = zc !== undefined && zc !== null && zp !== undefined && zp !== null;
            // mismo orden de operaciones que en Python: gradiente * (z_punto - z_comunidad)
            const ajuste = corregida ? gradiente * (zp - zc) : 0;
            const pm = propsMun[p.g] || {};
            const props = {
                GID_3: p.id, comunidad: p.n, municipio: pm.municipio || '', gid_municipio: p.g,
                departamento: pm.departamento || '', provincia: pm.provincia || '', poblacion: p.p,
                z: corregida ? zc : null, z_celda: zp ?? null, ajuste_temp: ajuste, corregida,
                celda: clave, lat: p.y, lon: p.x,
            };
            filaDe[p.id] = ids.length;
            propsDe[p.id] = props;
            ids.push(p.id);
            filas.push(i);
            ajustes.push(ajuste);
            features.push({ type: 'Feature', geometry: f.geometry, properties: props });
        });

        const D = {};
        Object.entries(PC.datos).forEach(([v, porModelo]) => {
            D[v] = {};
            Object.entries(porModelo).forEach(([m, mat]) => {
                D[v][m] = VARIABLES_TEMPERATURA.includes(v)
                    ? filas.map((i, j) => { const a = ajustes[j]; return mat[i].map((x) => (x === null ? null : x + a)); })
                    : filas.map((i) => mat[i]);   // la misma fila de la celda (sin copiar)
            });
        });
        const P = { ...PC, municipios: ids, datos: D };

        const CULT = {};
        Object.entries(CULTmun).forEach(([k, c]) => {
            const muns = c.municipios;
            const mun = {};
            ids.forEach((id) => {
                const pr = propsDe[id];
                if (muns[pr.gid_municipio]) mun[id] = { nombres: [pr.comunidad], comunidad: true, sup_ha: null, prod_tn: null };
            });
            CULT[k] = { ...c, municipios: mun, municipios_productores: muns };
        });

        const info = {
            resolucion: res,
            n_celdas: PC.municipios.length,
            n_comunidades: ids.length,
            con_altitud: Object.keys(z).length > 0,
            fuente_altitud: ELEV ? ELEV.fuente : null,
        };
        return { P, MUN: { type: 'FeatureCollection', features }, CULT, propsDe, filaDe, info };
    }

    function cargar(MUNmun, CULTmun) {
        if (!promesa) {
            promesa = construir(MUNmun, CULTmun).then((d) => { datos = d; return d; })
                .catch((e) => { promesa = null; throw e; });
        }
        return promesa;
    }

    // ------------------------------------------------------------------ alertas por comunidad
    // Mismos umbrales que los scripts de Python (PLANTILLAS del simulador); la condición «k»
    // del simulador es «cond_k» en alertas_latest.json.
    function indicadorComunidades(ind, clave, P, CULT) {
        const S = window.SimuladorRiesgo;
        const pl = S.PLANTILLAS[clave];
        const n = Math.min(pl.ventana, P.fechas.length);
        const activas = Object.entries(pl.conds);
        const idsCultivo = Object.keys((CULT[ind.clave_cultivo] || { municipios: {} }).municipios);
        const fila = datos.filaDe;
        const nuevo = {
            ...ind,
            ventana: { inicio: P.fechas[0], fin: P.fechas[n - 1], dias: n },
            municipios: idsCultivo,
        };
        let calculado = null;
        const calcular = () => {
            if (calculado) return calculado;
            const out = { indice: {}, alerta: {}, variables: {}, condiciones_cumplidas: {} };
            const vars = Object.keys(ind.variables);
            vars.forEach((v) => { out.variables[v] = {}; });
            activas.forEach(([k]) => { out.condiciones_cumplidas['cond_' + k] = {}; });
            P.modelos.forEach((m) => {
                const s = (v) => (P.datos[v] && P.datos[v][m] ? P.datos[v][m] : null);
                const tx = s('temperature_2m_max'), tn = s('temperature_2m_min');
                const pp = s('precipitation_sum'), hr = s('relative_humidity_2m');
                const filaDe = (mat, f) => (mat ? mat[f] : []);
                const r = idsCultivo.map((id) => {
                    const f = fila[id];
                    const vp = S.variablesPeriodo(filaDe(tx, f), filaDe(tn, f), filaDe(pp, f), filaDe(hr, f), n);
                    return { vp, ev: S.evaluar(vp, activas) };
                });
                out.indice[m] = r.map((x) => x.ev.indice);
                out.alerta[m] = r.map((x) => x.ev.alerta);
                vars.forEach((v) => { out.variables[v][m] = r.map((x) => x.vp[v] ?? null); });
                activas.forEach(([k]) => { out.condiciones_cumplidas['cond_' + k][m] = r.map((x) => x.ev.conds[k]); });
            });
            calculado = out;
            return out;
        };
        ['indice', 'alerta', 'variables', 'condiciones_cumplidas'].forEach((campo) => {
            Object.defineProperty(nuevo, campo, { get: () => calcular()[campo], enumerable: true, configurable: true });
        });
        return nuevo;
    }

    /** Alertas por comunidad con el formato de alertas_latest.json (se calculan al pedir cada indicador). */
    function alertas(Amun) {
        if (!datos) throw new Error('Primero hay que cargar los datos de comunidades');
        if (cacheAlertas.has(Amun)) return cacheAlertas.get(Amun);
        const { P, CULT } = datos;
        const indicadores = {};
        Object.entries(Amun.indicadores).forEach(([k, ind]) => {
            if (window.SimuladorRiesgo && window.SimuladorRiesgo.PLANTILLAS[k]) indicadores[k] = indicadorComunidades(ind, k, P, CULT);
        });
        const A = {
            ...Amun, indicadores, nivel: 'comunidades',
            modelos: P.modelos, fecha_inicio: P.fecha_inicio,
            generado_utc: P.generado_utc, pronostico_generado_utc: P.generado_utc,
        };
        cacheAlertas.set(Amun, A);
        return A;
    }

    function cambiarNivel(n) {
        if (n !== 'municipios' && n !== 'comunidades') return;
        if (n === nivel) return;
        nivel = n;
        document.dispatchEvent(new CustomEvent('sat:nivel', { detail: { nivel } }));
    }

    /** Texto con el municipio y departamento de una comunidad. */
    function lugar(p) {
        if (!p) return '';
        return [p.municipio, p.departamento].filter(Boolean).join(', ');
    }

    /** Frase sobre cómo se obtuvo el pronóstico de una comunidad. */
    function notaPronostico(p, info) {
        const res = info.resolucion ? `${String(info.resolucion).replace('.', ',')}°` : '';
        let t = `Pronóstico de la celda de ${res} que contiene la comunidad`;
        if (p.corregida) {
            const a = p.ajuste_temp;
            const signo = a > 0 ? '+' : a < 0 ? '−' : '±';
            t += `; temperatura ajustada ${signo}${Math.abs(a).toLocaleString('es-BO', { maximumFractionDigits: 1 })} °C ` +
                `por altitud (comunidad a ${p.z.toLocaleString('es-BO')} m, punto de la celda a ${Math.round(p.z_celda).toLocaleString('es-BO')} m).`;
        } else {
            t += ' (sin ajuste de temperatura por altitud).';
        }
        return t;
    }

    window.SATComunidades = {
        get nivel() { return nivel; },
        cambiarNivel, cargar, alertas, lugar, notaPronostico, claveCelda,
        get datos() { return datos; },
    };

    // Los selectores «Ver resultados por» de los tres exploradores quedan sincronizados
    document.addEventListener('DOMContentLoaded', () => {
        const selects = [...document.querySelectorAll('select[data-sat-nivel]')];
        selects.forEach((s) => {
            s.value = nivel;
            s.addEventListener('change', () => cambiarNivel(s.value));
        });
        // Textos fijos de la página con versión para comunidades (atributo data-com)
        const textos = () => {
            document.querySelectorAll('[data-com]').forEach((e) => {
                if (e.dataset.mun === undefined) e.dataset.mun = e.innerHTML;
                e.innerHTML = nivel === 'comunidades' ? e.dataset.com : e.dataset.mun;
            });
        };
        document.addEventListener('sat:nivel', () => { selects.forEach((s) => { s.value = nivel; }); textos(); });
    });
})();
