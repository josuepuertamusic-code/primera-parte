(function (global) {
  "use strict";

  /** Cantidad de decimales escritos en un número (máximo 4). */
  function decimalesDe(valor) {
    const texto = String(valor);
    if (/e-/i.test(texto)) return Math.min(4, Number(texto.split(/e-/i)[1]) || 0);
    return texto.includes(".") ? Math.min(4, texto.split(".")[1].length) : 0;
  }

  /** Redondeo "normal" (0,5 sube) evitando errores de coma flotante. */
  function redondear(valor, dec) {
    const f = 10 ** dec;
    return Math.sign(valor) * Math.round(Math.abs(valor) * f + 1e-9) / f;
  }

  function suma(arr) { return arr.reduce((a, b) => a + b, 0); }
  function promedio(arr) { return suma(arr) / arr.length; }

  /** PERCENTIL.INC de Excel: posición = 1 + (n − 1)·p, interpolando. */
  function percentil(ordenados, p) {
    const n = ordenados.length;
    if (n === 1) return { valor: ordenados[0], posicion: 1, bajo: ordenados[0], alto: ordenados[0], fraccion: 0 };
    const posicion = 1 + (n - 1) * p;
    const entero = Math.floor(posicion + 1e-9);
    const fraccion = Math.abs(posicion - entero) < 1e-9 ? 0 : posicion - entero;
    const bajo = ordenados[entero - 1];
    const alto = ordenados[Math.min(n - 1, entero)];
    return { valor: bajo + fraccion * (alto - bajo), posicion, bajo, alto, fraccion };
  }

  /** Moda de los datos sin agrupar (puede haber varias). */
  function modas(valores) {
    const conteo = new Map();
    valores.forEach(v => conteo.set(v, (conteo.get(v) || 0) + 1));
    const max = Math.max(...conteo.values());
    if (max <= 1) return { valores: [], repeticiones: 1 };
    const lista = [...conteo.entries()].filter(([, c]) => c === max).map(([v]) => v).sort((a, b) => a - b);
    return { valores: lista, repeticiones: max };
  }

  /* ------------------------------------------------------------------ */
  /*  Distribución de frecuencias completa                               */
  /* ------------------------------------------------------------------ */
  function calcular(valores) {
    if (!Array.isArray(valores) || valores.length < 2) return null;
    const n = valores.length;
    const dec = Math.max(0, ...valores.map(decimalesDe));
    const escala = 10 ** dec;
    const unidad = 1 / escala;
    // Trabajamos en "unidades enteras" para que 26,1 + 0,1 dé 26,2 exacto.
    const enteros = valores.map(v => Math.round(v * escala));
    const ordenados = [...valores].sort((a, b) => a - b);
    const minimo = ordenados[0];
    const maximo = ordenados[n - 1];
    const minE = Math.min(...enteros);
    const maxE = Math.max(...enteros);

    // PASO 1 · Sturges
    const kCrudo = 1 + 3.322 * Math.log10(n);
    const k = Math.ceil(redondear(kCrudo, 9));

    // PASO 2 · Longitud
    const rango = (maxE - minE) / escala;
    const longitudCruda = rango / k;
    const longitudE = Math.max(1, Math.ceil(redondear((maxE - minE) / k, 9)));
    const longitud = longitudE / escala;

    // PASO 3 · Intervalos  (mayor = menor + L − unidad)
    const intervalosE = [];
    let menorE = minE;
    const agregar = () => {
      const mayorE = menorE + longitudE - 1;
      intervalosE.push({ menorE, mayorE });
      menorE = mayorE + 1;
    };
    for (let i = 0; i < k; i++) agregar();
    let agregados = 0;
    while (maxE > intervalosE[intervalosE.length - 1].mayorE) { agregar(); agregados++; }

    // PASO 4 · Tabla de frecuencias
    let acumulado = 0;
    const clases = intervalosE.map((it, i) => {
      const fi = enteros.filter(v => v >= it.menorE && v <= it.mayorE).length;
      acumulado += fi;
      const menor = it.menorE / escala;
      const mayor = it.mayorE / escala;
      const mi = (it.menorE + it.mayorE) / 2 / escala;
      return {
        ci: i + 1, menor, mayor, fi, Fi: acumulado,
        hi: fi / n, Hi: acumulado / n, mi, fiMi: fi * mi
      };
    });

    // Medidas de tendencia central
    const media = promedio(valores);
    const mediaAgrupada = suma(clases.map(c => c.fiMi)) / n;
    const todosPositivos = valores.every(v => v > 0);
    const geometrica = todosPositivos ? Math.exp(promedio(valores.map(Math.log))) : null;
    const armonica = todosPositivos ? n / suma(valores.map(v => 1 / v)) : null;
    const moda = modas(valores);

    // Moda de Czuber (datos agrupados)
    const maxFi = Math.max(...clases.map(c => c.fi));
    const iModal = clases.findIndex(c => c.fi === maxFi);
    const modal = clases[iModal];
    const fiAnterior = iModal > 0 ? clases[iModal - 1].fi : 0;
    const fiSiguiente = iModal < clases.length - 1 ? clases[iModal + 1].fi : 0;
    const d1 = modal.fi - fiAnterior;
    const d2 = modal.fi - fiSiguiente;
    const czuber = d1 + d2 === 0 ? modal.mi : modal.menor + (d1 / (d1 + d2)) * longitud;
    const clasesModales = clases.filter(c => c.fi === maxFi).map(c => c.ci);

    // Medidas de posición (PERCENTIL.INC)
    const pos = {
      p10: percentil(ordenados, 0.10),
      q1: percentil(ordenados, 0.25),
      mediana: percentil(ordenados, 0.50),
      q3: percentil(ordenados, 0.75),
      p90: percentil(ordenados, 0.90)
    };

    // Dispersión
    const sumaCuadrados = suma(valores.map(v => (v - media) ** 2));
    const varianzaP = sumaCuadrados / n;
    const varianzaM = sumaCuadrados / (n - 1);
    const desviacionP = Math.sqrt(varianzaP);
    const desviacionM = Math.sqrt(varianzaM);
    const cvP = media !== 0 ? desviacionP / media : null;
    const cvM = media !== 0 ? desviacionM / media : null;
    const ric = pos.q3.valor - pos.q1.valor;
    const errorTipico = desviacionM / Math.sqrt(n);
    const desviacionMedia = suma(valores.map(v => Math.abs(v - media))) / n;

    // Forma (fórmulas de Excel COEFICIENTE.ASIMETRIA y CURTOSIS)
    let asimetria = null, curtosis = null;
    if (n >= 3 && desviacionM > 0) {
      const z3 = suma(valores.map(v => ((v - media) / desviacionM) ** 3));
      asimetria = (n / ((n - 1) * (n - 2))) * z3;
    }
    if (n >= 4 && desviacionM > 0) {
      const z4 = suma(valores.map(v => ((v - media) / desviacionM) ** 4));
      curtosis = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3)) * z4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    }
    // Asimetría de Pearson (fácil de explicar): 3·(media − mediana) / s
    const asimetriaPearson = desviacionM > 0 ? 3 * (media - pos.mediana.valor) / desviacionM : null;

    // Bigotes del diagrama de caja (regla 1,5 · RIC)
    const cercaInf = pos.q1.valor - 1.5 * ric;
    const cercaSup = pos.q3.valor + 1.5 * ric;
    const dentro = ordenados.filter(v => v >= cercaInf - 1e-9 && v <= cercaSup + 1e-9);
    const atipicos = ordenados.filter(v => v < cercaInf - 1e-9 || v > cercaSup + 1e-9);

    return {
      valores, ordenados, n, dec, unidad, minimo, maximo, suma: suma(valores),
      kCrudo, k, kFinal: clases.length, agregados,
      rango, longitudCruda, longitud, clases,
      media, mediaAgrupada, geometrica, armonica, moda,
      czuber: { valor: czuber, clase: modal, d1, d2, fiAnterior, fiSiguiente, empates: clasesModales },
      pos,
      sumaCuadrados, varianzaP, varianzaM, desviacionP, desviacionM, cvP, cvM, ric,
      errorTipico, desviacionMedia, asimetria, curtosis, asimetriaPearson,
      caja: { cercaInf, cercaSup, bigoteInf: dentro[0], bigoteSup: dentro[dentro.length - 1], atipicos },
      regresion: regresion(valores)
    };
  }

  /* ------------------------------------------------------------------ */
  /*  Percentil a pedido (cualquier k de 1 a 99)                         */
  /* ------------------------------------------------------------------ */
  function percentilK(valores, k) {
    const ord = [...valores].sort((a, b) => a - b);
    return percentil(ord, k / 100);
  }

  /* ------------------------------------------------------------------ */
  /*  Regresión: X = posición del dato (1, 2, 3…), Y = valor             */
  /* ------------------------------------------------------------------ */
  const NOMBRES_MODELO = {
    lineal: "Lineal", exponencial: "Exponencial", logaritmico: "Logarítmico",
    potencial: "Potencial", polinomico: "Polinómico (grado 2)"
  };

  function predecir(m, x) {
    const [a, b, c = 0] = m.coef;
    switch (m.id) {
      case "lineal": return a + b * x;
      case "exponencial": return a * Math.exp(b * x);
      case "logaritmico": return a + b * Math.log(x);
      case "potencial": return a * x ** b;
      case "polinomico": return a + b * x + c * x * x;
      default: return NaN;
    }
  }

  function minimosCuadrados(x, y) {
    const mx = promedio(x), my = promedio(y);
    let sxy = 0, sxx = 0;
    for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; }
    const b = sxy / sxx;
    return [my - b * mx, b];
  }

  function resolver3(A, B) {
    const M = A.map((fila, i) => [...fila, B[i]]);
    for (let c = 0; c < 3; c++) {
      let p = c;
      for (let r = c + 1; r < 3; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
      [M[c], M[p]] = [M[p], M[c]];
      for (let r = c + 1; r < 3; r++) {
        const f = M[r][c] / M[c][c];
        for (let j = c; j < 4; j++) M[r][j] -= f * M[c][j];
      }
    }
    const s = [0, 0, 0];
    for (let r = 2; r >= 0; r--) {
      let t = M[r][3];
      for (let j = r + 1; j < 3; j++) t -= M[r][j] * s[j];
      s[r] = t / M[r][r];
    }
    return s;
  }

  function polinomio2(x, y) {
    // Se centra y escala X para evitar errores numéricos con muchos datos.
    const mx = promedio(x);
    const esc = Math.max(1, (x[x.length - 1] - x[0]) / 2);
    const S = [0, 0, 0, 0, 0], T = [0, 0, 0];
    for (let i = 0; i < x.length; i++) {
      const u = (x[i] - mx) / esc;
      let p = 1;
      for (let e = 0; e <= 4; e++) { S[e] += p; if (e <= 2) T[e] += p * y[i]; p *= u; }
    }
    const [a0, a1, a2] = resolver3([[S[0], S[1], S[2]], [S[1], S[2], S[3]], [S[2], S[3], S[4]]], T);
    const b = a1 / esc, c = a2 / (esc * esc);
    return [a0 - b * mx + c * mx * mx, b - 2 * c * mx, c];
  }

  function evaluar(id, coef, x, y, p) {
    if (!coef.every(Number.isFinite)) return { id, nombre: NOMBRES_MODELO[id], aplica: false, motivo: "No se pudo ajustar con estos datos." };
    const m = { id, coef };
    const my = promedio(y);
    let sse = 0, sst = 0;
    for (let i = 0; i < x.length; i++) { sse += (y[i] - predecir(m, x[i])) ** 2; sst += (y[i] - my) ** 2; }
    const n = x.length;
    const r2 = sst === 0 ? 1 : 1 - sse / sst;
    const gl = n - p;
    const r2a = gl > 0 ? 1 - (1 - r2) * (n - 1) / gl : NaN;
    return { id, nombre: NOMBRES_MODELO[id], coef, r2, r2a, aplica: true };
  }

  function regresion(y) {
    const n = y.length;
    const x = y.map((_, i) => i + 1);
    const positivos = y.every(v => v > 0);
    const noAplica = (id, motivo) => ({ id, nombre: NOMBRES_MODELO[id], aplica: false, motivo });
    const modelos = [
      evaluar("lineal", minimosCuadrados(x, y), x, y, 2),
      positivos ? (() => { const [a, b] = minimosCuadrados(x, y.map(Math.log)); return evaluar("exponencial", [Math.exp(a), b], x, y, 2); })()
                : noAplica("exponencial", "Necesita que todos los valores sean mayores que 0."),
      evaluar("logaritmico", minimosCuadrados(x.map(Math.log), y), x, y, 2),
      positivos ? (() => { const [a, b] = minimosCuadrados(x.map(Math.log), y.map(Math.log)); return evaluar("potencial", [Math.exp(a), b], x, y, 2); })()
                : noAplica("potencial", "Necesita que todos los valores sean mayores que 0."),
      n >= 4 ? evaluar("polinomico", polinomio2(x, y), x, y, 3) : noAplica("polinomico", "Necesita al menos 4 datos.")
    ];
    let mejor = null;
    modelos.forEach(m => { if (m.aplica && Number.isFinite(m.r2a) && (!mejor || m.r2a > mejor.r2a + 1e-4)) mejor = m; });
    // Correlación de Pearson entre posición y valor
    const mx = promedio(x), my = promedio(y);
    let sxy = 0, sxx = 0, syy = 0;
    for (let i = 0; i < n; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; }
    const r = syy === 0 ? 0 : sxy / Math.sqrt(sxx * syy);
    return { modelos, mejor, r };
  }

  /* ------------------------------------------------------------------ */
  /*  Promedios especiales (calculadoras pequeñas)                       */
  /* ------------------------------------------------------------------ */
  function ponderado(filas) {           // filas: [{peso, valor}]
    const sp = suma(filas.map(f => f.peso));
    const spx = suma(filas.map(f => f.peso * f.valor));
    return { sumaPesos: sp, sumaProductos: spx, resultado: sp ? spx / sp : null };
  }
  function geometricaDe(valores) {
    if (!valores.length || valores.some(v => v <= 0)) return null;
    return Math.exp(promedio(valores.map(Math.log)));
  }
  function armonicaDe(valores) {
    if (!valores.length || valores.some(v => v <= 0)) return null;
    return valores.length / suma(valores.map(v => 1 / v));
  }

  const API = { calcular, percentilK, predecir, ponderado, geometricaDe, armonicaDe, decimalesDe, redondear };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  global.Estadistica = API;
})(typeof window !== "undefined" ? window : globalThis);
