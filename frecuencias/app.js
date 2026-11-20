/* ==========================================================================
   app.js — Parte visual: lee los datos, muestra tablas, explicaciones,
   gráficos y descarga el Excel. Los cálculos están en estadistica.js.
   ========================================================================== */
(function () {
  "use strict";

  const E = window.Estadistica;
  const $ = id => document.getElementById(id);

  const EJEMPLOS = {
    unidades: {
      titulo: "Unidades vendidas",
      datos: [150,80,210,95,340,120,450,60,175,230,110,290,140,310,75,190,420,130,260,85,380,160,220,100,300,145,70,205,350,115,410,90,185,240,125,275,65,330,155,215,105,390,135,250,80,320,170,200,440,95,165,235,110,285,75,315,140,205,375,125,430,85,190,260,100,345,150,225,70,295,120,360,180,210,90,400,130,245,105,325,160,220,65,280,115,355,175,200,450,95,145,230,110,305,80,335,185,215,125,395]
    },
    precio: {
      titulo: "Precio unitario promedio (USD)",
      datos: [45.5,62.1,38,55.4,25.8,50,19.5,75.3,42.1,35.6,48.9,29.4,46.2,28,68.5,40.3,21,49.7,32.4,59.8,23.5,44.1,36.9,53.2,30.1,45.8,70.4,39.2,24.6,51.5,20.8,57,41.6,34.2,47.9,31.5,72.8,26.4,43.7,37.4,52.6,22.2,48.1,33.8,61.5,27.3,42.9,40,19.8,54.9,43.2,35.1,49.3,30.8,67.2,27.9,46.5,39.7,23.9,48.4,20.2,60.5,40.8,32.1,53.8,25.1,45.1,36.4,71,29.7,50.6,24.2,41.9,38.5,58.1,21.5,49.1,34.7,52.3,26.9,44.6,37.1,73.5,31.2,51.9,24.8,42.4,39.9,19.1,55.7,46.8,35.9,49.8,29.2,62.8,25.6,41.3,37.8,48.7,21.9]
    }
  };

  let R = null;          // último resultado
  let graficos = {};     // instancias de Chart.js

  const PALETA = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300",
                  "#6250d6", "#e34948", "#0091a8", "#b06a2c", "#9b2f8f", "#4f8a10"];
  const COLOR_LINEA = "#334155";
  const COLOR_RESALTE = "#eb6834";
  const COLOR_TEXTO = "#172033";
  const COLOR_TEXTO_2 = "#52596a";
  const COLOR_GRILLA = "#e6eaf0";

  function hexARgb(h) { const n = parseInt(h.slice(1), 16); return [n >> 16 & 255, n >> 8 & 255, n & 255]; }
  function rgbAHex(c) { return "#" + c.map(v => Math.round(v).toString(16).padStart(2, "0")).join(""); }
  function coloresIntervalos(k) {
    return Array.from({ length: k }, (_, i) => PALETA[i % PALETA.length]);
  }
  /** Texto oscuro sobre colores claros (amarillo, verde, rosado) y blanco sobre los oscuros. */
  function textoSobre(hex) {
    const [r, g, b] = hexARgb(hex).map(v => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; });
    return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.3 ? COLOR_TEXTO : "#ffffff";
  }

  /* ------------------------------------------------------------------ */
  /*  Formato de números (coma decimal, como en Excel en español)        */
  /* ------------------------------------------------------------------ */
  function dec() { return Number($("decimales").value); }
  function f(v, d = dec()) {
    if (v === null || v === undefined || !Number.isFinite(v)) return "—";
    return v.toLocaleString("es-CO", { minimumFractionDigits: d, maximumFractionDigits: d });
  }
  function fl(v) { return f(v, R ? R.dec : 0); }                 // límites con la precisión de los datos
  function fx(v) { return f(v, R ? R.dec : 0); }                 // datos crudos con sus propios decimales
  function pct(h, d = dec()) { return Number.isFinite(h) ? f(h * 100, d) + " %" : "—"; }
  function esc(t) { return String(t).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function titulo() { return $("titulo").value.trim() || "Valores"; }

  function tabla(cabeceras, filas, pie = null, clases = []) {
    const th = cabeceras.map((h, i) => `<th class="${clases[i] || ""}">${h}</th>`).join("");
    const tb = filas.map(fila => `<tr>${fila.map((c, i) => `<td class="${clases[i] || ""}">${c}</td>`).join("")}</tr>`).join("");
    const tf = pie ? `<tfoot><tr>${pie.map((c, i) => `<td class="${clases[i] || ""}">${c}</td>`).join("")}</tr></tfoot>` : "";
    return `<div class="tabla-wrap"><table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody>${tf}</table></div>`;
  }

  /* ------------------------------------------------------------------ */
  /*  Lectura de los datos escritos                                      */
  /* ------------------------------------------------------------------ */
  function aNumero(texto, comaEsDecimal) {
    let t = String(texto).trim().replace(/\s+/g, "").replace(/^"|"$/g, "");
    if (t === "") return NaN;
    if (comaEsDecimal) t = t.replace(",", ".");
    return /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t) ? Number(t) : NaN;
  }

  function leerTexto(texto) {
    const valores = [], errores = [];
    let encabezado = null;
    let lineas = texto.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (lineas.length === 1 && /[;\t ]/.test(lineas[0])) lineas = lineas[0].split(/[;\t ]+/).filter(Boolean);
    // Si la primera línea es un encabezado con comas (CSV "N,Valor"), la coma separa columnas.
    const comaSepara = lineas.length > 1 && lineas[0].includes(",") && !Number.isFinite(aNumero(lineas[0], true)) &&
      lineas.slice(1, 6).every(l => l.split(",").length === lineas[0].split(",").length);
    lineas.forEach((linea, i) => {
      let celdas;
      if (/[;\t]/.test(linea)) celdas = linea.split(/[;\t]/);
      else if (comaSepara) celdas = linea.split(",");
      else celdas = [linea];
      celdas = celdas.map(c => c.trim()).filter(c => c !== "");
      const ultimo = celdas[celdas.length - 1] ?? "";
      const v = aNumero(ultimo, !comaSepara);
      if (Number.isFinite(v)) { valores.push(v); return; }
      if (i === 0 && valores.length === 0) { encabezado = ultimo; return; }
      errores.push(`Línea ${i + 1}: «${linea.slice(0, 40)}» no es un número.`);
    });
    return { valores, errores, encabezado };
  }

  function actualizarConteo() {
    const { valores, errores } = leerTexto($("entrada").value);
    $("conteo").textContent = valores.length >= 2
      ? `${valores.length.toLocaleString("es-CO")} números listos para calcular.`
      : `${valores.length} números. Se necesitan al menos 2.`;
    const ul = $("errores");
    ul.hidden = !errores.length;
    ul.innerHTML = errores.slice(0, 8).map(e => `<li>${esc(e)}</li>`).join("") +
      (errores.length > 8 ? `<li>Y ${errores.length - 8} errores más.</li>` : "") +
      (errores.length ? `<li>Estas líneas no se tendrán en cuenta.</li>` : "");
  }

  /* ------------------------------------------------------------------ */
  /*  Cálculo principal                                                  */
  /* ------------------------------------------------------------------ */
  function calcularTodo(desplazar = true) {
    const { valores, encabezado } = leerTexto($("entrada").value);
    if (encabezado && ($("titulo").value.trim() === "" || $("titulo").dataset.auto === "1")) {
      $("titulo").value = encabezado; $("titulo").dataset.auto = "1";
    }
    if (valores.length < 2) {
      $("resultados").hidden = true;
      $("conteo").textContent = "Escribe al menos 2 números para poder calcular.";
      return;
    }
    R = E.calcular(valores);
    $("resultados").hidden = false;
    pintarResumen();
    pintarPasos();
    pintarTabla();
    pintarGraficos();
    pintarCentral();
    pintarPosicion();
    pintarDispersion();
    pintarForma();
    pintarRegresion();
    pintarGlosario();
    pintarPromedios();
    if (desplazar) $("resumen").scrollIntoView({ behavior: "smooth" });
  }

  /* ---------------------------- 2. Resumen ---------------------------- */
  function pintarResumen() {
    const t = titulo();
    const tarjetas = [
      ["Total de datos (n)", R.n.toLocaleString("es-CO"), "Cuántos números escribiste."],
      ["Mínimo", fx(R.minimo), "El dato más pequeño."],
      ["Máximo", fx(R.maximo), "El dato más grande."],
      ["Media (promedio)", f(R.media), "Suma de todos ÷ cantidad de datos."],
      ["Mediana", f(R.pos.mediana.valor), "El valor del medio: la mitad está por debajo."],
      ["Desviación estándar", f(R.desviacionP), "Cuánto se alejan, en promedio, los datos de la media."]
    ];
    $("tarjetas").innerHTML = tarjetas.map(([a, b, c]) => `<div class="tarjeta"><span>${a}</span><strong>${b}</strong><small>${c}</small></div>`).join("");

    const dif = R.media - R.pos.mediana.valor;
    const relDif = R.desviacionP ? Math.abs(dif) / R.desviacionP : 0;
    let comparacion;
    if (relDif < 0.1) comparacion = "La media y la mediana son casi iguales, así que los datos están bastante equilibrados alrededor del centro.";
    else if (dif > 0) comparacion = `La media (${f(R.media)}) es mayor que la mediana (${f(R.pos.mediana.valor)}): hay algunos valores altos que “jalan” el promedio hacia arriba.`;
    else comparacion = `La media (${f(R.media)}) es menor que la mediana (${f(R.pos.mediana.valor)}): hay algunos valores bajos que “jalan” el promedio hacia abajo.`;
    $("resumenTexto").innerHTML = `
      <p><strong>En palabras sencillas:</strong> analizaste <strong>${R.n}</strong> datos de <em>${esc(t)}</em>.
      El valor más bajo fue <strong>${fx(R.minimo)}</strong> y el más alto <strong>${fx(R.maximo)}</strong> (una diferencia, o rango, de ${fl(R.rango)}).</p>
      <p>En promedio, el valor es <strong>${f(R.media)}</strong>. La mitad de los datos está por debajo de <strong>${f(R.pos.mediana.valor)}</strong> y la otra mitad por encima.
      El 50 % de los datos del centro está entre <strong>${f(R.pos.q1.valor)}</strong> y <strong>${f(R.pos.q3.valor)}</strong>.</p>
      <p>${comparacion}</p>`;
  }

  /* ---------------------------- 3. Pasos ------------------------------ */
  function fila(etq, contenido) { return `<div class="formula-fila"><span>${etq}</span><span>${contenido}</span></div>`; }

  function pintarPasos() {
    const u = R.unidad;
    const log = Math.log10(R.n);
    const paso1 = `
      <div class="paso">
        <h3>Paso 1 · Número de intervalos (k) — Regla de Sturges</h3>
        ${fila("Fórmula", "<code>k = 1 + 3,322 × log₁₀(n)</code>")}
        ${fila("Reemplazando", `<code>k = 1 + 3,322 × log₁₀(${R.n}) = 1 + 3,322 × ${f(log, 4)}</code>`)}
        ${fila("Resultado", `<code>k = ${f(R.kCrudo, 3)}</code>`)}
        ${fila("Se redondea hacia arriba", `<span class="resultado">k = ${R.k} intervalos</span>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> La regla de Sturges sugiere en cuántos grupos (intervalos) repartir los datos según cuántos hay. Con ${R.n} datos conviene usar ${R.k} grupos. Se redondea siempre hacia arriba (en Excel: REDONDEAR.MAS) para que no falten grupos.</p>
      </div>`;
    const paso2 = `
      <div class="paso">
        <h3>Paso 2 · Longitud o ancho de cada intervalo (L)</h3>
        ${fila("Fórmula", "<code>L = (dato mayor − dato menor) ÷ k</code>")}
        ${fila("Reemplazando", `<code>L = (${fx(R.maximo)} − ${fx(R.minimo)}) ÷ ${R.k} = ${fl(R.rango)} ÷ ${R.k}</code>`)}
        ${fila("Resultado", `<code>L = ${f(R.longitudCruda, Math.max(3, R.dec + 2))}</code>`)}
        ${fila("Se redondea hacia arriba", `<span class="resultado">L = ${fl(R.longitud)}</span>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> Es el tamaño de cada grupo. Se reparte la distancia entre el dato más pequeño y el más grande en ${R.k} partes iguales. Se redondea hacia arriba con ${R.dec === 0 ? "números enteros" : R.dec + " decimal(es)"} (los mismos decimales de tus datos) para que el último intervalo alcance a incluir al dato mayor.</p>
      </div>`;
    const filasInt = R.clases.map(c => [c.ci, fl(c.menor), fl(c.mayor)]);
    const paso3 = `
      <div class="paso">
        <h3>Paso 3 · Construir los intervalos</h3>
        ${fila("Primer menor", `<code>menor₁ = dato menor = ${fx(R.minimo)}</code>`)}
        ${fila("Cada mayor", `<code>mayor = menor + L − ${fl(u)}</code>`)}
        ${fila("Siguiente menor", `<code>menor siguiente = mayor anterior + ${fl(u)}</code>`)}
        ${fila("Ejemplo", `<code>${fl(R.clases[0].menor)} + ${fl(R.longitud)} − ${fl(u)} = ${fl(R.clases[0].mayor)}</code>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> Se resta ${fl(u)} (la unidad más pequeña de tus datos) para que los intervalos no se “pisen”: un dato no puede quedar en dos grupos a la vez.
        ${R.agregados ? `<br><strong>Nota:</strong> el dato mayor (${fx(R.maximo)}) no cabía en el intervalo ${R.k}, por eso se agregó ${R.agregados === 1 ? "un intervalo más" : R.agregados + " intervalos más"}.` : ""}</p>
      </div>`;
    const paso3t = `<div class="paso">
        <h3>Intervalos obtenidos</h3>
        ${tabla(["Intervalo", "Menor", "Mayor"], filasInt, null, ["", "n", "n"])}
      </div>`;
    const paso4 = `
      <div class="paso paso-ancho">
        <h3>Paso 4 · Contar cuántos datos caen en cada intervalo</h3>
        <p>Se revisa cada dato y se cuenta en qué intervalo cae (fi). Con esos conteos se calculan las demás columnas. El resultado es la <a href="#tabla">tabla de distribución de frecuencias (sección 4)</a>.</p>
      </div>`;
    $("pasosContenido").innerHTML = paso1 + paso2 + paso3 + paso3t + paso4;
  }

  /* ---------------------------- 4. Tabla ------------------------------ */
  function pintarTabla() {
    const colores = coloresIntervalos(R.clases.length);
    const filas = R.clases.map((c, i) => [
      `<span class="punto-color" style="background:${colores[i]}"></span>`,
      c.ci, fl(c.menor), fl(c.mayor), c.fi, c.Fi, f(c.hi, Math.max(2, dec())), f(c.Hi, Math.max(2, dec())), pct(c.hi), f(c.mi), f(c.fiMi)
    ]);
    const sumaFiMi = R.clases.reduce((a, c) => a + c.fiMi, 0);
    $("tablaFrecuencias").innerHTML = tabla(
      ["", "Ci", "Menor", "Mayor", "fi", "Fi", "hi = fi ÷ n", "Hi = Fi ÷ n", "Porcentaje", "Mi", "fi · Mi"],
      filas,
      ["", "Total", "", "", R.n, "", f(1, Math.max(2, dec())), "", "100 %", "", f(sumaFiMi)],
      ["color-celda", "", "n", "n", "n", "n", "n", "n", "n", "n", "n"]
    );
    const mod = R.czuber.clase;
    const iMitad = R.clases.findIndex(c => c.Hi >= 0.5);
    const cMitad = R.clases[iMitad];
    const menosFrec = R.clases.reduce((a, c) => c.fi < a.fi ? c : a, R.clases[0]);
    $("tablaInterpretacion").innerHTML = `
      <p><strong>Lo que dice la tabla:</strong> el intervalo con más datos es el <strong>${mod.ci}</strong> (de ${fl(mod.menor)} a ${fl(mod.mayor)}), con <strong>${mod.fi}</strong> datos, es decir el <strong>${pct(mod.hi)}</strong> del total.
      ${R.czuber.empates.length > 1 ? `(Hay empate: los intervalos ${R.czuber.empates.join(", ")} tienen la misma cantidad).` : ""}</p>
      <p>El intervalo con menos datos es el <strong>${menosFrec.ci}</strong> (${fl(menosFrec.menor)} a ${fl(menosFrec.mayor)}) con ${menosFrec.fi} dato(s).</p>
      <p>Leyendo la columna Hi: hasta el intervalo ${cMitad.ci} (valores de hasta ${fl(cMitad.mayor)}) ya se acumula el <strong>${pct(cMitad.Hi)}</strong> de los datos.</p>`;
  }

  /* ---------------------------- 5. Gráficos --------------------------- */
  const etiquetasPlugin = {
    id: "etiquetasValor",
    afterDatasetsDraw(chart, _args, opts) {
      if (!opts || !opts.activo) return;
      const { ctx } = chart;
      ctx.save();
      ctx.font = "600 12px Segoe UI, Roboto, Arial";
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        meta.data.forEach((el, i) => {
          const texto = opts.texto(ds, i);
          if (texto === null || texto === undefined || texto === "") return;
          if (chart.config.type === "pie") {
            const p = el.tooltipPosition();
            ctx.fillStyle = textoSobre(ds.backgroundColor[i]);
            ctx.textAlign = "center"; ctx.textBaseline = "middle";
            ctx.fillText(texto, p.x, p.y);
          } else {
            ctx.fillStyle = COLOR_TEXTO;
            ctx.textAlign = "center"; ctx.textBaseline = "bottom";
            ctx.fillText(texto, el.x, el.y - 6);
          }
        });
      });
      ctx.restore();
    }
  };

  function destruir() { Object.values(graficos).forEach(g => g.destroy()); graficos = {}; }

  function ejes(xTitulo, yTitulo, extraY = {}) {
    return {
      x: { grid: { display: false }, ticks: { color: COLOR_TEXTO_2 }, title: { display: true, text: xTitulo, color: COLOR_TEXTO_2 } },
      y: { beginAtZero: true, grid: { color: COLOR_GRILLA }, ticks: { color: COLOR_TEXTO_2 }, title: { display: true, text: yTitulo, color: COLOR_TEXTO_2 }, ...extraY }
    };
  }
  const tooltipBase = { backgroundColor: COLOR_TEXTO, titleColor: "#fff", bodyColor: "#fff", padding: 10, cornerRadius: 6 };

  function pintarGraficos() {
    const colores = coloresIntervalos(R.clases.length);
    const etiquetas = R.clases.map(c => `${fl(c.menor)} – ${fl(c.mayor)}`);
    $("leyenda").innerHTML = "<strong style='font-size:.86rem;margin-right:6px'>Intervalos:</strong>" +
      R.clases.map((c, i) => `<span class="leyenda-item"><span class="punto-color" style="background:${colores[i]}"></span>${c.ci}. ${etiquetas[i]}</span>`).join("");

    if (!window.Chart) {
      ["gBarras", "gPastel", "gPoligono", "gOjiva", "gRegresion"].forEach(id => {
        $(id).parentElement.innerHTML = "<p class='nota'>No se pudo cargar la librería de gráficos (revisa tu conexión a internet).</p>";
      });
      pintarCaja();
      return;
    }
    destruir();
    Chart.defaults.font.family = "Segoe UI, Roboto, Arial, sans-serif";
    const muchos = R.clases.length > 12;

    // Histograma
    graficos.barras = new Chart($("gBarras"), {
      type: "bar",
      data: { labels: etiquetas, datasets: [{ label: "fi", data: R.clases.map(c => c.fi), backgroundColor: colores, borderRadius: { topLeft: 4, topRight: 4 }, borderSkipped: "bottom", categoryPercentage: 1, barPercentage: 0.96 }] },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 20 } },
        plugins: {
          legend: { display: false },
          tooltip: { ...tooltipBase, callbacks: { title: it => `Intervalo ${it[0].dataIndex + 1}: ${it[0].label}`, label: it => `${it.raw} datos (${pct(R.clases[it.dataIndex].hi)})` } },
          etiquetasValor: { activo: !muchos, texto: (ds, i) => ds.data[i] }
        },
        scales: ejes(titulo(), "Cantidad de datos (fi)", { ticks: { precision: 0, color: COLOR_TEXTO_2 } })
      },
      plugins: [etiquetasPlugin]
    });

    // Pastel
    graficos.pastel = new Chart($("gPastel"), {
      type: "pie",
      data: { labels: etiquetas, datasets: [{ data: R.clases.map(c => c.hi * 100), backgroundColor: colores, borderColor: "#fff", borderWidth: 2 }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { ...tooltipBase, callbacks: { title: it => `Intervalo ${it[0].dataIndex + 1}: ${it[0].label}`, label: it => `${pct(R.clases[it.dataIndex].hi)} de los datos (${R.clases[it.dataIndex].fi} datos)` } },
          etiquetasValor: { activo: true, texto: (ds, i) => R.clases[i].hi >= 0.05 ? f(R.clases[i].hi * 100, 0) + "%" : "" }
        }
      },
      plugins: [etiquetasPlugin]
    });

    // Polígono (con punto en cero antes y después, como en Excel)
    const L = R.longitud;
    const puntosPol = [{ x: R.clases[0].mi - L, y: 0 }, ...R.clases.map(c => ({ x: c.mi, y: c.fi })), { x: R.clases[R.clases.length - 1].mi + L, y: 0 }];
    const coloresPol = ["#ffffff", ...colores, "#ffffff"];
    graficos.poligono = new Chart($("gPoligono"), {
      type: "line",
      data: { datasets: [{ label: "fi", data: puntosPol, borderColor: COLOR_LINEA, borderWidth: 2, pointBackgroundColor: coloresPol, pointBorderColor: COLOR_LINEA, pointBorderWidth: 1.5, pointRadius: 6, pointHoverRadius: 8, tension: 0, fill: { target: "origin", above: "rgba(51,65,85,0.06)" } }] },
      options: {
        responsive: true, maintainAspectRatio: false, parsing: false, layout: { padding: { top: 20 } },
        plugins: {
          legend: { display: false },
          tooltip: { ...tooltipBase, callbacks: {
            title: it => { const i = it[0].dataIndex - 1; return i >= 0 && i < R.clases.length ? `Intervalo ${i + 1}: ${etiquetas[i]}` : "Punto de cierre"; },
            label: it => `Mi = ${f(it.raw.x)} · fi = ${it.raw.y}` } },
          etiquetasValor: { activo: !muchos, texto: (ds, i) => (i === 0 || i === ds.data.length - 1) ? "" : ds.data[i].y }
        },
        scales: {
          x: { type: "linear", grid: { color: COLOR_GRILLA }, ticks: { color: COLOR_TEXTO_2, callback: v => f(v, R.dec) }, title: { display: true, text: `${titulo()} (marca de clase Mi)`, color: COLOR_TEXTO_2 } },
          y: { beginAtZero: true, grid: { color: COLOR_GRILLA }, ticks: { precision: 0, color: COLOR_TEXTO_2 }, title: { display: true, text: "Cantidad de datos (fi)", color: COLOR_TEXTO_2 } }
        }
      },
      plugins: [etiquetasPlugin]
    });

    // Ojiva (Mi contra Hi, como en el Excel de clase)
    graficos.ojiva = new Chart($("gOjiva"), {
      type: "line",
      data: { datasets: [{ label: "Hi", data: R.clases.map(c => ({ x: c.mi, y: c.Hi * 100 })), borderColor: COLOR_LINEA, borderWidth: 2, pointBackgroundColor: colores, pointBorderColor: COLOR_LINEA, pointBorderWidth: 1.5, pointRadius: 6, pointHoverRadius: 8, tension: 0, fill: false }] },
      options: {
        responsive: true, maintainAspectRatio: false, parsing: false, layout: { padding: { top: 20 } },
        plugins: {
          legend: { display: false },
          tooltip: { ...tooltipBase, callbacks: { title: it => `Intervalo ${it[0].dataIndex + 1}: ${etiquetas[it[0].dataIndex]}`, label: it => `Hasta aquí: ${pct(R.clases[it.dataIndex].Hi)} de los datos (${R.clases[it.dataIndex].Fi} datos)` } },
          etiquetasValor: { activo: !muchos, texto: (ds, i) => f(ds.data[i].y, 0) + "%" }
        },
        scales: {
          x: { type: "linear", grid: { color: COLOR_GRILLA }, ticks: { color: COLOR_TEXTO_2, callback: v => f(v, R.dec) }, title: { display: true, text: `${titulo()} (marca de clase Mi)`, color: COLOR_TEXTO_2 } },
          y: { min: 0, max: 110, grid: { color: COLOR_GRILLA }, ticks: { stepSize: 10, color: COLOR_TEXTO_2, callback: v => v <= 100 ? v + "%" : "" }, title: { display: true, text: "Porcentaje acumulado (Hi)", color: COLOR_TEXTO_2 } }
        }
      },
      plugins: [etiquetasPlugin]
    });

    const mod = R.czuber.clase;
    $("txtBarras").innerHTML = `<br><strong>Con tus datos:</strong> la barra más alta es la del intervalo ${mod.ci} (${fl(mod.menor)} – ${fl(mod.mayor)}) con ${mod.fi} datos.`;
    $("txtPastel").innerHTML = `<br><strong>Con tus datos:</strong> la porción más grande es el intervalo ${mod.ci}, con el ${pct(mod.hi)} de los datos.`;
    $("txtPoligono").innerHTML = `<br><strong>Con tus datos:</strong> el punto más alto (la “cima”) está en Mi = ${f(mod.mi)}.`;
    const c50 = R.clases.find(c => c.Hi >= 0.5);
    $("txtOjiva").innerHTML = `<br><strong>Con tus datos:</strong> la línea pasa el 50 % en el intervalo ${c50.ci} (${fl(c50.menor)} – ${fl(c50.mayor)}); la mediana exacta es ${f(R.pos.mediana.valor)}.`;
    pintarCaja();
  }

  function pintarCaja() {
    const W = 900, H = 170, m = { l: 30, r: 30 };
    const lo = Math.min(R.minimo, R.caja.cercaInf < R.minimo ? R.minimo : R.minimo);
    const hi = R.maximo;
    const span = hi - lo || 1;
    const x = v => m.l + (v - lo) / span * (W - m.l - m.r);
    const y0 = 40, h = 56, yc = y0 + h / 2;
    const { q1, mediana, q3 } = R.pos;
    const ticks = Array.from({ length: 6 }, (_, i) => lo + span * i / 5);
    const atip = R.caja.atipicos.map(v => `<circle cx="${x(v)}" cy="${yc}" r="5" fill="#fff" stroke="${COLOR_RESALTE}" stroke-width="2"><title>Atípico: ${fx(v)}</title></circle>`).join("");
    const etiqueta = (v, txt, dy) => `<text x="${x(v)}" y="${dy}" text-anchor="middle" font-size="12" fill="${COLOR_TEXTO}">${txt}</text>`;
    $("gCaja").innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" font-family="Segoe UI, Arial, sans-serif" role="img" aria-label="Diagrama de caja: Q1 ${f(q1.valor)}, mediana ${f(mediana.valor)}, Q3 ${f(q3.valor)}">
        ${ticks.map(t => `<line x1="${x(t)}" x2="${x(t)}" y1="20" y2="${y0 + h + 14}" stroke="${COLOR_GRILLA}"/><text x="${x(t)}" y="${H - 12}" text-anchor="middle" font-size="12" fill="${COLOR_TEXTO_2}">${f(t, R.dec)}</text>`).join("")}
        <line x1="${x(R.caja.bigoteInf)}" x2="${x(q1.valor)}" y1="${yc}" y2="${yc}" stroke="${COLOR_LINEA}" stroke-width="2"/>
        <line x1="${x(q3.valor)}" x2="${x(R.caja.bigoteSup)}" y1="${yc}" y2="${yc}" stroke="${COLOR_LINEA}" stroke-width="2"/>
        <line x1="${x(R.caja.bigoteInf)}" x2="${x(R.caja.bigoteInf)}" y1="${yc - 14}" y2="${yc + 14}" stroke="${COLOR_LINEA}" stroke-width="2"/>
        <line x1="${x(R.caja.bigoteSup)}" x2="${x(R.caja.bigoteSup)}" y1="${yc - 14}" y2="${yc + 14}" stroke="${COLOR_LINEA}" stroke-width="2"/>
        <rect x="${x(q1.valor)}" y="${y0}" width="${Math.max(1, x(q3.valor) - x(q1.valor))}" height="${h}" rx="4" fill="#cde2fb" stroke="#2a78d6" stroke-width="2"><title>Caja: de Q1 = ${f(q1.valor)} a Q3 = ${f(q3.valor)}</title></rect>
        <line x1="${x(mediana.valor)}" x2="${x(mediana.valor)}" y1="${y0}" y2="${y0 + h}" stroke="${COLOR_RESALTE}" stroke-width="3"><title>Mediana: ${f(mediana.valor)}</title></line>
        <path d="M ${x(R.media)} ${yc - 6} l 6 6 l -6 6 l -6 -6 z" fill="${COLOR_TEXTO}"><title>Media: ${f(R.media)}</title></path>
        ${atip}
        ${etiqueta(q1.valor, "Q1 " + f(q1.valor), y0 - 8)}
        ${etiqueta(q3.valor, "Q3 " + f(q3.valor), y0 - 8)}
        ${etiqueta(mediana.valor, "Mediana " + f(mediana.valor), y0 + h + 30)}
      </svg>
      <p class="nota" style="margin:4px 0 0"><span style="color:${COLOR_RESALTE};font-weight:700">│</span> mediana &nbsp; ◆ media &nbsp; ${R.caja.atipicos.length ? `<span style="color:${COLOR_RESALTE}">○</span> valores atípicos` : ""}</p>`;
    $("txtCaja").innerHTML = `<br><strong>Con tus datos:</strong> el 50 % central está entre ${f(q1.valor)} y ${f(q3.valor)}. Los bigotes llegan de ${fx(R.caja.bigoteInf)} a ${fx(R.caja.bigoteSup)}. ` +
      (R.caja.atipicos.length ? `Hay ${R.caja.atipicos.length} valor(es) atípico(s): ${R.caja.atipicos.map(fx).join(", ")}.` : "No hay valores atípicos (ningún dato está demasiado lejos del resto).") +
      ` Se considera atípico un dato menor que Q1 − 1,5 × RIC = ${f(R.caja.cercaInf)} o mayor que Q3 + 1,5 × RIC = ${f(R.caja.cercaSup)}.`;
  }

  /* ---------------------------- 6. Tendencia central ------------------ */
  function pintarCentral() {
    const n = R.n;
    const cz = R.czuber;
    const modaTxt = R.moda.valores.length
      ? `${R.moda.valores.map(fx).join("; ")} <br><span class="etiqueta info">${R.moda.valores.length === 1 ? "Unimodal" : R.moda.valores.length === 2 ? "Bimodal" : "Multimodal"}</span>`
      : `No existe <br><span class="etiqueta info">Amodal</span>`;
    const filas = [
      ["<strong>Media aritmética</strong>", "x̄ = Σx ÷ n", `${f(R.suma)} ÷ ${n}`, f(R.media),
        "El promedio común: si repartiéramos el total en partes iguales, a cada dato le tocaría esto. Se usa cuando todos los datos tienen la misma importancia."],
      ["<strong>Media con datos agrupados</strong>", "x̄ ≈ Σ(fi · Mi) ÷ n", `${f(R.mediaAgrupada * n)} ÷ ${n}`, f(R.mediaAgrupada),
        "Es el promedio usando la tabla (cada dato se reemplaza por la marca de clase de su intervalo). Da un valor aproximado al promedio real."],
      ["<strong>Mediana</strong>", "Dato del centro (posición (n + 1) ÷ 2)", R.n % 2 ? `Posición ${(n + 1) / 2}: dato central` : `Posiciones ${n / 2} y ${n / 2 + 1}: (${fx(R.ordenados[n / 2 - 1])} + ${fx(R.ordenados[n / 2])}) ÷ 2`, f(R.pos.mediana.valor),
        "Ordenando los datos de menor a mayor, la mitad queda por debajo y la mitad por encima. No se deja afectar por valores muy altos o muy bajos."],
      ["<strong>Moda (datos sin agrupar)</strong>", "Dato que más se repite", R.moda.valores.length ? `Se repite ${R.moda.repeticiones} veces` : "Ningún dato se repite", modaTxt,
        "El valor más frecuente. Puede no existir (amodal) o haber varias (bimodal, multimodal)."],
      ["<strong>Moda con datos agrupados (Czuber)</strong>", "Mo = Li + [d₁ ÷ (d₁ + d₂)] × L",
        `d₁ = ${cz.clase.fi} − ${cz.fiAnterior} = ${cz.d1}; d₂ = ${cz.clase.fi} − ${cz.fiSiguiente} = ${cz.d2}<br>${fl(cz.clase.menor)} + [${cz.d1} ÷ (${cz.d1} + ${cz.d2})] × ${fl(R.longitud)}`, f(cz.valor),
        `Estima el valor más frecuente usando la tabla. Li es el límite menor del intervalo con más datos (intervalo ${cz.clase.ci}); d₁ y d₂ comparan su fi con el intervalo anterior y el siguiente.`],
      ["<strong>Media geométrica</strong>", "G = ⁿ√(x₁ × x₂ × … × xₙ)", R.geometrica === null ? "Necesita que todos los datos sean mayores que 0" : `${n}√(producto de los ${n} datos)`, f(R.geometrica),
        "Se usa con tasas, porcentajes o crecimientos. Siempre es menor o igual que la media aritmética."],
      ["<strong>Media armónica</strong>", "H = n ÷ Σ(1 ÷ xᵢ)", R.armonica === null ? "Necesita que todos los datos sean mayores que 0" : `${n} ÷ ${f(n / R.armonica, 6)}`, f(R.armonica),
        "Se usa con razones o velocidades (por ejemplo km/h). Siempre es menor o igual que la geométrica."]
    ];
    $("tablaCentral").innerHTML = tabla(["Medida", "Fórmula", "Reemplazando", "Resultado", "¿Qué significa?"], filas, null, ["", "formula-celda", "formula-celda", "valor", "significado"]);
    $("centralTexto").innerHTML = `<p><strong>¿Cuál representa mejor a tus datos?</strong> Si hay valores muy extremos, la <strong>mediana</strong> (${f(R.pos.mediana.valor)}) es más confiable que la media (${f(R.media)}), porque no se deja “jalar” por esos valores.
      ${R.geometrica !== null ? `Fíjate que se cumple: armónica (${f(R.armonica)}) ≤ geométrica (${f(R.geometrica)}) ≤ aritmética (${f(R.media)}).` : ""}</p>`;
  }

  /* ---------------------------- 7. Posición --------------------------- */
  function descPos(p) {
    if (p.fraccion === 0) return `Posición ${f(p.posicion, 0)} → dato ${fx(p.bajo)}`;
    return `Posición ${f(p.posicion, 2)} → ${fx(p.bajo)} + ${f(p.fraccion, 2)} × (${fx(p.alto)} − ${fx(p.bajo)})`;
  }
  function pintarPosicion() {
    const t = titulo();
    const filas = [
      ["Mínimo", "—", "Primer dato ordenado", fx(R.minimo), "Ningún dato es menor."],
      ["Percentil 10 (P10)", `1 + (${R.n} − 1) × 0,10`, descPos(R.pos.p10), f(R.pos.p10.valor), `El 10 % de los datos es menor o igual a ${f(R.pos.p10.valor)}.`],
      ["Cuartil 1 (Q1 = P25)", `1 + (${R.n} − 1) × 0,25`, descPos(R.pos.q1), f(R.pos.q1.valor), `La cuarta parte (25 %) de los datos es menor o igual a ${f(R.pos.q1.valor)}.`],
      ["Mediana (Q2 = P50)", `1 + (${R.n} − 1) × 0,50`, descPos(R.pos.mediana), f(R.pos.mediana.valor), `La mitad de los datos es menor o igual a ${f(R.pos.mediana.valor)}.`],
      ["Cuartil 3 (Q3 = P75)", `1 + (${R.n} − 1) × 0,75`, descPos(R.pos.q3), f(R.pos.q3.valor), `Tres cuartas partes (75 %) de los datos son menores o iguales a ${f(R.pos.q3.valor)}.`],
      ["Percentil 90 (P90)", `1 + (${R.n} − 1) × 0,90`, descPos(R.pos.p90), f(R.pos.p90.valor), `Solo el 10 % de los datos supera ${f(R.pos.p90.valor)}.`],
      ["Máximo", "—", "Último dato ordenado", fx(R.maximo), "Ningún dato es mayor."]
    ];
    $("tablaPosicion").innerHTML = tabla(["Medida", "Posición = 1 + (n − 1) × p", "Cálculo", "Resultado", "¿Qué significa?"], filas, null, ["", "formula-celda", "formula-celda", "valor", "significado"]);
    $("ordenados").innerHTML = R.ordenados.map((v, i) => `<span data-i="${i}"><b>${i + 1}</b>${fx(v)}</span>`).join("");
    calcPk();
  }
  function calcPk() {
    if (!R) return;
    const k = Math.min(99, Math.max(1, Math.round(Number($("pk").value) || 50)));
    $("pk").value = k;
    const p = E.percentilK(R.valores, k);
    document.querySelectorAll("#ordenados span").forEach(s => s.classList.remove("marcado"));
    const iBajo = Math.floor(p.posicion + 1e-9) - 1;
    [iBajo, p.fraccion ? iBajo + 1 : iBajo].forEach(i => { const s = document.querySelector(`#ordenados span[data-i="${i}"]`); if (s) s.classList.add("marcado"); });
    $("resPk").innerHTML = `<p><code>posición = 1 + (${R.n} − 1) × ${f(k / 100, 2)} = ${f(p.posicion, 2)}</code></p>
      <p>${descPos(p)} = <strong>P${k} = ${f(p.valor)}</strong></p>
      <p><strong>Significa:</strong> el ${k} % de los datos es menor o igual a ${f(p.valor)}, y el ${100 - k} % es mayor. (Los datos usados están resaltados en la lista de abajo).</p>`;
  }

  /* ---------------------------- 8. Dispersión ------------------------- */
  function nivelCV(cv) {
    if (cv === null) return ["—", "info"];
    const v = Math.abs(cv) * 100;
    if (v < 15) return ["Datos homogéneos (poca variación)", "ok"];
    if (v <= 30) return ["Variación moderada", "aviso"];
    return ["Datos heterogéneos (mucha variación)", "alto"];
  }
  function pintarDispersion() {
    const n = R.n;
    const [cvTxt, cvClase] = nivelCV(R.cvP);
    const filas = [
      ["<strong>Rango</strong>", "R = máximo − mínimo", `${fx(R.maximo)} − ${fx(R.minimo)}`, fl(R.rango), "La distancia entre el dato más pequeño y el más grande."],
      ["<strong>Varianza poblacional</strong>", "σ² = Σ(x − x̄)² ÷ n", `${f(R.sumaCuadrados)} ÷ ${n}`, f(R.varianzaP), "El promedio de las distancias al cuadrado entre cada dato y la media. Está en unidades al cuadrado, por eso se usa más la desviación."],
      ["<strong>Desviación estándar poblacional</strong>", "σ = √σ²", `√${f(R.varianzaP)}`, f(R.desviacionP), `En promedio, los datos se alejan ${f(R.desviacionP)} de la media (${f(R.media)}).`],
      ["<strong>Varianza muestral</strong>", "s² = Σ(x − x̄)² ÷ (n − 1)", `${f(R.sumaCuadrados)} ÷ ${n - 1}`, f(R.varianzaM), "Igual que la anterior, pero dividiendo entre n − 1 porque los datos son una muestra."],
      ["<strong>Desviación estándar muestral</strong>", "s = √s²", `√${f(R.varianzaM)}`, f(R.desviacionM), "La desviación estándar cuando tus datos son una muestra."],
      ["<strong>Coeficiente de variación</strong>", "CV = (σ ÷ x̄) × 100", `(${f(R.desviacionP)} ÷ ${f(R.media)}) × 100`, pct(R.cvP), `Compara la desviación con el promedio, en porcentaje. <span class="etiqueta ${cvClase}">${cvTxt}</span>`],
      ["<strong>Rango intercuartílico (RIC)</strong>", "RIC = Q3 − Q1", `${f(R.pos.q3.valor)} − ${f(R.pos.q1.valor)}`, f(R.ric), "El ancho del 50 % central de los datos (la caja del diagrama de caja)."],
      ["<strong>Desviación media</strong>", "DM = Σ|x − x̄| ÷ n", `${f(R.desviacionMedia * n)} ÷ ${n}`, f(R.desviacionMedia), "El promedio de las distancias (sin elevar al cuadrado) entre cada dato y la media."],
      ["<strong>Error típico</strong>", "EE = s ÷ √n", `${f(R.desviacionM)} ÷ √${n}`, f(R.errorTipico), "Qué tanto podría cambiar el promedio si se tomara otra muestra del mismo tamaño."]
    ];
    $("tablaDispersion").innerHTML = tabla(["Medida", "Fórmula", "Reemplazando", "Resultado", "¿Qué significa?"], filas, null, ["", "formula-celda", "formula-celda", "valor", "significado"]);
    $("dispersionTexto").innerHTML = `<p><strong>En palabras sencillas:</strong> el promedio es ${f(R.media)} y los datos se alejan de él, en promedio, unos <strong>${f(R.desviacionP)}</strong>.
      Eso es el <strong>${pct(R.cvP)}</strong> del promedio, por lo que se considera: <span class="etiqueta ${cvClase}">${cvTxt}</span>.</p>
      <p>Guía para el coeficiente de variación: menos de 15 % = datos homogéneos (parecidos entre sí); de 15 % a 30 % = variación moderada; más de 30 % = datos heterogéneos (muy diferentes entre sí).</p>`;
  }

  /* ---------------------------- 9. Forma ------------------------------ */
  function pintarForma() {
    const a = R.asimetria, k = R.curtosis;
    let tipoA = "—", txtA = "Se necesitan al menos 3 datos distintos.", claseA = 1;
    if (a !== null) {
      if (Math.abs(a) < 0.5) { tipoA = "Aproximadamente simétrica"; claseA = 1; txtA = "Los datos se reparten de forma parecida a ambos lados del centro."; }
      else if (a > 0) { tipoA = "Asimetría positiva (sesgo a la derecha)"; claseA = 2; txtA = "La mayoría de los datos está en los valores bajos y hay una “cola” de pocos valores altos."; }
      else { tipoA = "Asimetría negativa (sesgo a la izquierda)"; claseA = 0; txtA = "La mayoría de los datos está en los valores altos y hay una “cola” de pocos valores bajos."; }
    }
    let tipoK = "—", txtK = "Se necesitan al menos 4 datos distintos.";
    if (k !== null) {
      if (k > 0.5) { tipoK = "Leptocúrtica (puntiaguda)"; txtK = "Muchos datos se amontonan cerca del centro: la forma tiene un pico alto."; }
      else if (k < -0.5) { tipoK = "Platicúrtica (aplanada)"; txtK = "Los datos están repartidos de forma pareja: la forma es más plana que una campana."; }
      else { tipoK = "Mesocúrtica (normal)"; txtK = "La altura del pico es parecida a la de una campana normal."; }
    }
    const filas = [
      ["<strong>Coeficiente de asimetría</strong> (Excel: COEFICIENTE.ASIMETRIA)", "As = n ÷ [(n − 1)(n − 2)] × Σ[(x − x̄) ÷ s]³", f(a, Math.max(dec(), 3)), `<span class="etiqueta info">${tipoA}</span><br>${txtA}`],
      ["<strong>Asimetría de Pearson</strong>", "As = 3 × (x̄ − Mediana) ÷ s", f(R.asimetriaPearson, Math.max(dec(), 3)), "Otra forma más sencilla de medir la asimetría: compara la media con la mediana. Mismo criterio: cerca de 0 = simétrica."],
      ["<strong>Curtosis</strong> (Excel: CURTOSIS)", "K = n(n + 1) ÷ [(n − 1)(n − 2)(n − 3)] × Σ[(x − x̄) ÷ s]⁴ − 3(n − 1)² ÷ [(n − 2)(n − 3)]", f(k, Math.max(dec(), 3)), `<span class="etiqueta info">${tipoK}</span><br>${txtK}`]
    ];
    $("tablaForma").innerHTML = tabla(["Medida", "Fórmula", "Resultado", "¿Qué significa?"], filas, null, ["", "formula-celda", "valor", "significado"]);
    $("formaTexto").innerHTML = `<p><strong>Cómo interpretar:</strong> asimetría entre −0,5 y 0,5 = simétrica; mayor que 0,5 = cola hacia la derecha; menor que −0,5 = cola hacia la izquierda.
      Curtosis mayor que 0,5 = puntiaguda; entre −0,5 y 0,5 = normal; menor que −0,5 = aplanada.</p>
      <p><strong>Tus datos:</strong> ${tipoA.toLowerCase()} y ${tipoK.toLowerCase()}. Compáralo con el histograma de la sección 5.</p>`;
    const curva = d => `<svg viewBox="0 0 120 60" aria-hidden="true"><path d="${d}" fill="#b7d3f6" stroke="${COLOR_LINEA}" stroke-width="2"/></svg>`;
    const formas = [
      ["Sesgo a la izquierda", "M5 55 C 40 55, 60 50, 80 10 C 92 -5, 105 40, 115 55 Z"],
      ["Simétrica", "M5 55 C 35 55, 45 8, 60 8 C 75 8, 85 55, 115 55 Z"],
      ["Sesgo a la derecha", "M5 55 C 15 40, 28 -5, 40 10 C 60 50, 80 55, 115 55 Z"]
    ];
    $("formasIlustracion").innerHTML = formas.map(([t, d], i) => `<div class="forma-mini ${a !== null && i === claseA ? "activa" : ""}">${curva(d)}<div>${t}${a !== null && i === claseA ? " ← tus datos" : ""}</div></div>`).join("");
  }

  /* ---------------------------- 10. Regresión ------------------------- */
  function ecuacion(m) {
    const [a, b, c] = m.coef.map(v => f(v, 4));
    const sg = v => (v < 0 ? " − " : " + ") + f(Math.abs(v), 4);   // "+ −0,03" → "− 0,03"
    const [, cb, cc] = m.coef;
    switch (m.id) {
      case "lineal": return `Y = ${a}${sg(cb)}·X`;
      case "exponencial": return `Y = ${a}·e^(${b}·X)`;
      case "logaritmico": return `Y = ${a}${sg(cb)}·ln(X)`;
      case "potencial": return `Y = ${a}·X^(${b})`;
      case "polinomico": return `Y = ${a}${sg(cb)}·X${sg(cc)}·X²`;
      default: return "";
    }
  }
  function pintarRegresion() {
    const rg = R.regresion;
    const filas = rg.modelos.map(m => m.aplica
      ? [m.nombre + (rg.mejor && rg.mejor.id === m.id ? ' <span class="etiqueta ok">Elegido</span>' : ""), `<code>${ecuacion(m)}</code>`, f(m.r2, 4), f(m.r2a, 4)]
      : [m.nombre, `<span class="nota">${m.motivo}</span>`, "—", "—"]);
    $("tablaRegresion").innerHTML = tabla(["Modelo", "Ecuación", "R²", "R² ajustado"], filas, null, ["", "", "valor", "valor"]);

    const mejor = rg.mejor;
    if (window.Chart && mejor) {
      const n = R.n;
      const paso = Math.max(1, Math.ceil(n / 1500));
      const puntos = [];
      for (let i = 0; i < n; i += paso) puntos.push({ x: i + 1, y: R.valores[i] });
      const curva = Array.from({ length: 121 }, (_, j) => { const xx = 1 + (n - 1) * j / 120; return { x: xx, y: E.predecir(mejor, xx) }; });
      graficos.regresion = new Chart($("gRegresion"), {
        type: "line",
        data: { datasets: [
          { label: "Tus datos", data: puntos, showLine: false, pointRadius: n > 300 ? 2 : 4, pointBackgroundColor: "rgba(42,120,214,0.55)", pointBorderColor: "#2a78d6", borderColor: "#2a78d6", backgroundColor: "#2a78d6" },
          { label: `Modelo ${mejor.nombre.toLowerCase()}`, data: curva, pointRadius: 0, borderColor: COLOR_RESALTE, backgroundColor: COLOR_RESALTE, borderWidth: 3, tension: 0 }
        ] },
        options: {
          responsive: true, maintainAspectRatio: false, parsing: false,
          interaction: { mode: "nearest", intersect: false },
          plugins: { legend: { position: "bottom", labels: { color: COLOR_TEXTO } }, tooltip: { ...tooltipBase, callbacks: { title: it => `Posición ${f(it[0].raw.x, 0)}`, label: it => `${it.dataset.label}: ${f(it.raw.y)}` } } },
          scales: {
            x: { type: "linear", min: 1, max: Math.max(2, n), grid: { color: COLOR_GRILLA }, ticks: { color: COLOR_TEXTO_2, precision: 0 }, title: { display: true, text: "X · posición del dato", color: COLOR_TEXTO_2 } },
            y: { grid: { color: COLOR_GRILLA }, ticks: { color: COLOR_TEXTO_2 }, title: { display: true, text: `Y · ${titulo()}`, color: COLOR_TEXTO_2 } }
          }
        }
      });
    }
    const r = rg.r;
    const fuerza = Math.abs(r) < 0.1 ? "prácticamente nula" : Math.abs(r) < 0.3 ? "débil" : Math.abs(r) < 0.7 ? "moderada" : "fuerte";
    const dir = Math.abs(r) < 0.1 ? "no hay una tendencia clara a subir ni a bajar" : r > 0 ? "los valores tienden a subir a medida que avanza el orden" : "los valores tienden a bajar a medida que avanza el orden";
    let conclusion;
    if (!mejor) conclusion = "No se pudo ajustar ningún modelo con estos datos.";
    else if (mejor.r2 < 0.1) conclusion = `El “mejor” modelo es el ${mejor.nombre.toLowerCase()}, pero su R² es apenas ${f(mejor.r2, 4)} (casi 0). <strong>Ningún modelo describe bien tus datos:</strong> los valores suben y bajan sin seguir un patrón según el orden en que se escribieron. Esto es normal si el orden no tiene un significado (por ejemplo, datos que no están en orden de tiempo).`;
    else conclusion = `El modelo que mejor se ajusta es el <strong>${mejor.nombre.toLowerCase()}</strong>: <code>${ecuacion(mejor)}</code>. Explica el <strong>${pct(mejor.r2)}</strong> de la variación de los datos (R² = ${f(mejor.r2, 4)}).`;
    $("regresionTexto").innerHTML = `<p>${conclusion}</p><p><strong>Correlación de Pearson:</strong> r = ${f(r, 4)} → relación ${fuerza}: ${dir}.</p>`;
  }

  /* ---------------------------- 11. Los cuatro promedios con tus datos */
  /*  Todo se calcula con los mismos valores escritos en la sección 1.   */
  function pintarPromedios() {
    const n = R.n;
    const v = R.valores;
    const positivos = v.every(x => x > 0);
    const sumaLog = positivos ? v.reduce((a, x) => a + Math.log10(x), 0) : null;
    const sumaInv = positivos ? v.reduce((a, x) => a + 1 / x, 0) : null;
    const sumaFi = R.clases.reduce((a, c) => a + c.fi, 0);
    const sumaFiMi = R.clases.reduce((a, c) => a + c.fiMi, 0);
    const ponderada = sumaFiMi / sumaFi;
    const noAplica = "No se puede calcular: todos los datos deben ser mayores que 0 (hay ceros o negativos).";

    const comparacion = tabla(["Promedio", "Resultado con tus datos", "¿Cuándo se usa?"], [
      ["<strong>a. Aritmético</strong>", f(R.media), "Cuando todos los datos tienen la misma importancia. Es el promedio común."],
      ["<strong>b. Ponderado</strong> (con la tabla)", f(ponderada), "Cuando cada valor tiene un peso distinto. Aquí cada marca de clase (Mi) pesa según cuántos datos tiene su intervalo (fi)."],
      ["<strong>c. Geométrico</strong>", positivos ? f(R.geometrica) : "No aplica", "Cuando los datos son tasas, porcentajes o crecimientos (se multiplican entre sí)."],
      ["<strong>d. Armónico</strong>", positivos ? f(R.armonica) : "No aplica", "Cuando los datos son razones o velocidades (por ejemplo km/h o unidades por hora)."]
    ], null, ["", "valor", "significado"]);

    const aritmetica = `
      <div class="paso">
        <h3>a. Promedio aritmético</h3>
        ${fila("Fórmula", "<code>x̄ = Σx ÷ n</code>")}
        ${fila("Suma de los datos", `<code>Σx = ${fx(v[0])} + ${fx(v[1])}${n > 2 ? " + … + " + fx(v[n - 1]) : ""} = ${f(R.suma)}</code>`)}
        ${fila("Reemplazando", `<code>x̄ = ${f(R.suma)} ÷ ${n}</code>`)}
        ${fila("Resultado", `<span class="resultado">x̄ = ${f(R.media)}</span>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> Si se repartiera el total (${f(R.suma)}) en partes iguales entre los ${n} datos, a cada uno le tocaría ${f(R.media)}.</p>
      </div>`;

    const ponderadoHtml = `
      <div class="paso">
        <h3>b. Promedio ponderado (usando la tabla de frecuencias)</h3>
        ${fila("Fórmula", "<code>x̄p = Σ(Xi · Pi) ÷ ΣPi</code>")}
        ${fila("Con tu tabla", "<code>Xi = marca de clase Mi · Pi = frecuencia fi</code>")}
        ${fila("Reemplazando", `<code>x̄p = ${f(sumaFiMi)} ÷ ${sumaFi}</code>`)}
        ${fila("Resultado", `<span class="resultado">x̄p = ${f(ponderada)}</span>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> Cada intervalo se representa con su punto medio (Mi), y ese punto “pesa” tanto como datos tenga el intervalo (fi). Por eso los intervalos con más datos influyen más. Da un valor muy cercano al promedio aritmético (${f(R.media)}).</p>
        ${tabla(["Ci", "Intervalo", "Xi = Mi", "Pi = fi", "Xi · Pi"],
          R.clases.map(c => [c.ci, `${fl(c.menor)} – ${fl(c.mayor)}`, f(c.mi), c.fi, f(c.fiMi)]),
          ["Total", "", "", sumaFi, f(sumaFiMi)], ["", "", "n", "n", "n"])}
      </div>`;

    const geometrica = positivos ? `
      <div class="paso">
        <h3>c. Media geométrica</h3>
        ${fila("Fórmula", "<code>G = ⁿ√(x₁ × x₂ × … × xₙ)</code>")}
        ${fila("Con logaritmos", "<code>G = 10 ^ (Σ log x ÷ n)</code>")}
        ${fila("Suma de logaritmos", `<code>Σ log x = log ${fx(v[0])} + log ${fx(v[1])}${n > 2 ? " + … + log " + fx(v[n - 1]) : ""} = ${f(sumaLog, 4)}</code>`)}
        ${fila("Reemplazando", `<code>G = 10 ^ (${f(sumaLog, 4)} ÷ ${n}) = 10 ^ ${f(sumaLog / n, 4)}</code>`)}
        ${fila("Resultado", `<span class="resultado">G = ${f(R.geometrica)}</span>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> Multiplicar ${n} números da un resultado gigante, por eso se usan logaritmos: es lo mismo que sacar la raíz ${n} del producto de todos los datos. La media geométrica (${f(R.geometrica)}) siempre es menor o igual que la aritmética (${f(R.media)}).</p>
      </div>` : `<div class="paso"><h3>c. Media geométrica</h3><p class="que-significa">${noAplica}</p></div>`;

    const armonica = positivos ? `
      <div class="paso">
        <h3>d. Media armónica</h3>
        ${fila("Fórmula", "<code>H = n ÷ Σ(1 ÷ x)</code>")}
        ${fila("Suma de inversos", `<code>Σ(1 ÷ x) = 1 ÷ ${fx(v[0])} + 1 ÷ ${fx(v[1])}${n > 2 ? " + … + 1 ÷ " + fx(v[n - 1]) : ""} = ${f(sumaInv, 6)}</code>`)}
        ${fila("Reemplazando", `<code>H = ${n} ÷ ${f(sumaInv, 6)}</code>`)}
        ${fila("Resultado", `<span class="resultado">H = ${f(R.armonica)}</span>`)}
        <p class="que-significa"><strong>¿Qué significa?</strong> Se divide la cantidad de datos entre la suma de sus inversos (1 ÷ cada dato). Da más importancia a los valores pequeños. Siempre es menor o igual que la geométrica (${f(R.geometrica)}).</p>
      </div>` : `<div class="paso"><h3>d. Media armónica</h3><p class="que-significa">${noAplica}</p></div>`;

    const detalle = `
      <details class="ayuda">
        <summary>Ver el cálculo dato por dato (de dónde sale cada suma)</summary>
        ${tabla(["N°", titulo(), "log x", "1 ÷ x"],
          v.map((x, i) => [i + 1, fx(x), positivos ? f(Math.log10(x), 4) : "—", positivos ? f(1 / x, 6) : "—"]),
          ["Total (Σ)", f(R.suma), positivos ? f(sumaLog, 4) : "—", positivos ? f(sumaInv, 6) : "—"], ["", "n", "n", "n"])}
      </details>`;

    const conclusion = positivos
      ? `<p><strong>Comparando:</strong> armónica (${f(R.armonica)}) ≤ geométrica (${f(R.geometrica)}) ≤ aritmética (${f(R.media)}). Esto siempre se cumple. Mientras más separados estén los datos, más diferentes salen estos tres promedios.</p>
         <p>Para datos como <em>${esc(titulo())}</em>, donde cada dato tiene la misma importancia, el promedio que se usa normalmente es el <strong>aritmético (${f(R.media)})</strong>.</p>`
      : `<p>Como hay datos en cero o negativos, solo se pueden usar el promedio aritmético (${f(R.media)}) y el ponderado (${f(ponderada)}).</p>`;

    $("promediosContenido").innerHTML = comparacion +
      `<div class="pasos" style="margin-top:14px">${aritmetica}${ponderadoHtml}${geometrica}${armonica}</div>` +
      detalle + `<div class="explicacion">${conclusion}</div>`;
  }

  /* ---------------------------- 12. Glosario -------------------------- */
  const GLOSARIO = [
    ["n", "n = cantidad de datos", "El total de números analizados."],
    ["Regla de Sturges", "k = 1 + 3,322 × log₁₀(n)", "Cuántos intervalos usar. Se redondea hacia arriba."],
    ["Rango", "R = máximo − mínimo", "Distancia entre el dato mayor y el menor."],
    ["Longitud del intervalo", "L = R ÷ k", "Ancho de cada intervalo. Se redondea hacia arriba a los decimales de los datos."],
    ["Límites del intervalo", "mayor = menor + L − unidad", "La unidad es 1 si los datos son enteros, 0,1 si tienen un decimal, etc."],
    ["Marca de clase", "Mi = (menor + mayor) ÷ 2", "El punto medio del intervalo."],
    ["Frecuencia absoluta", "fi = cuántos datos hay en el intervalo", "Conteo directo."],
    ["Frecuencia acumulada", "Fi = f₁ + f₂ + … + fi", "Datos desde el primer intervalo hasta el actual."],
    ["Frecuencia relativa", "hi = fi ÷ n", "Proporción del total. × 100 = porcentaje."],
    ["Frecuencia relativa acumulada", "Hi = Fi ÷ n", "Proporción acumulada. La última es 1 (100 %)."],
    ["Media aritmética", "x̄ = Σx ÷ n", "El promedio común."],
    ["Media agrupada", "x̄ ≈ Σ(fi · Mi) ÷ n", "Promedio calculado con la tabla."],
    ["Media ponderada", "x̄p = Σ(Xi · Pi) ÷ ΣPi", "Promedio cuando cada dato tiene distinta importancia."],
    ["Media geométrica", "G = ⁿ√(x₁ · x₂ · … · xₙ)", "Para tasas y porcentajes pequeños."],
    ["Media armónica", "H = n ÷ Σ(1 ÷ xᵢ)", "Para razones, velocidades y porcentajes de cambio."],
    ["Mediana", "Dato de la posición (n + 1) ÷ 2", "Divide los datos ordenados en dos mitades."],
    ["Moda", "Dato que más se repite", "Amodal (ninguna), unimodal (una), bimodal (dos)…"],
    ["Moda de Czuber", "Mo = Li + [d₁ ÷ (d₁ + d₂)] × L", "Moda estimada con datos agrupados."],
    ["Percentil (PERCENTIL.INC)", "posición = 1 + (n − 1) × p", "Valor que deja por debajo al p % de los datos."],
    ["Cuartiles", "Q1 = P25 · Q2 = P50 · Q3 = P75", "Parten los datos en cuatro grupos de 25 %."],
    ["Rango intercuartílico", "RIC = Q3 − Q1", "Ancho del 50 % central."],
    ["Varianza poblacional", "σ² = Σ(x − x̄)² ÷ n", "Excel: VARP / VAR.P."],
    ["Varianza muestral", "s² = Σ(x − x̄)² ÷ (n − 1)", "Excel: VAR.S."],
    ["Desviación estándar", "σ = √σ²   ·   s = √s²", "Distancia promedio de los datos a la media."],
    ["Coeficiente de variación", "CV = (σ ÷ x̄) × 100", "< 15 % homogéneo · 15–30 % moderado · > 30 % heterogéneo."],
    ["Error típico", "EE = s ÷ √n", "Precisión del promedio."],
    ["Asimetría", "As = n ÷ [(n−1)(n−2)] × Σ[(x − x̄) ÷ s]³", "0 = simétrica; + cola a la derecha; − cola a la izquierda."],
    ["Curtosis", "K = n(n+1) ÷ [(n−1)(n−2)(n−3)] × Σ[(x − x̄) ÷ s]⁴ − 3(n−1)² ÷ [(n−2)(n−3)]", "+ puntiaguda · 0 normal · − aplanada."],
    ["Valores atípicos", "x < Q1 − 1,5·RIC  o  x > Q3 + 1,5·RIC", "Datos muy alejados del resto."],
    ["Coeficiente de determinación", "R² = 1 − Σ(y − ŷ)² ÷ Σ(y − ȳ)²", "Qué parte de la variación explica el modelo (0 a 1)."],
    ["R² ajustado", "R²a = 1 − (1 − R²)(n − 1) ÷ (n − p)", "R² penalizado por la cantidad de coeficientes p."],
    ["Correlación de Pearson", "r = Σ(x − x̄)(y − ȳ) ÷ √[Σ(x − x̄)² · Σ(y − ȳ)²]", "−1 a 1: dirección y fuerza de la tendencia."]
  ];
  function pintarGlosario() {
    $("glosarioContenido").innerHTML = GLOSARIO.map(([t, fo, d]) => `<div class="formula-item"><strong>${t}</strong><code>${fo}</code><p>${d}</p></div>`).join("");
  }

  /* ---------------------------- Excel --------------------------------- */
  /*  Recorre cada sección de la página y la copia a una hoja de Excel:   */
  /*  títulos, textos, fórmulas, explicaciones, tablas (con los colores  */
  /*  de cada intervalo) y los gráficos como imágenes.                   */
  /* -------------------------------------------------------------------- */
  const COLS_TEXTO = 8;                       // columnas que ocupa un párrafo
  const BORDE = { style: "thin", color: { argb: "FFDBE1EA" } };
  const BORDES = { top: BORDE, left: BORDE, bottom: BORDE, right: BORDE };
  const argb = hex => "FF" + hex.replace("#", "").toUpperCase();
  const rellenar = hex => ({ type: "pattern", pattern: "solid", fgColor: { argb: argb(hex) } });

  function textoDe(el) {
    return (el.innerText || el.textContent || "").replace(/ /g, " ").replace(/[ \t]+/g, " ").replace(/\s*\n\s*/g, "\n").trim();
  }
  function cssAHex(css) { const m = String(css).match(/\d+(\.\d+)?/g); return m ? rgbAHex(m.slice(0, 3).map(Number)) : "#ffffff"; }
  function numeroEs(s) { const v = Number(s.replace(/\./g, "").replace(",", ".")); return Number.isFinite(v) ? v : null; }
  function formatoDec(s) { const p = s.split(",")[1]; return p ? "." + "0".repeat(p.length) : ""; }
  /** Convierte "1.234,56" en número y "21,00 %" en porcentaje, para que Excel pueda operar con ellos. */
  function valorCelda(texto) {
    const t = texto.trim().replace(/−/g, "-");
    const m = t.match(/^(-?\d[\d.]*(?:,\d+)?)\s?%$/);
    if (m && numeroEs(m[1]) !== null) return { v: numeroEs(m[1]) / 100, fmt: "0" + formatoDec(m[1]) + "%" };
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$|^-?\d+(,\d+)?$/.test(t)) return { v: numeroEs(t), fmt: "#,##0" + formatoDec(t) };
    return { v: texto.trim() };
  }

  class Hoja {
    constructor(libro, nombre) {
      this.libro = libro;
      this.ws = libro.addWorksheet(nombre.slice(0, 31), { views: [{ showGridLines: false }] });
      this.fila = 1;
      this.anchos = Array(30).fill(0);
      this.parrafos = [];     // [fila, texto, desdeCol]
      this.celdas = [];       // [fila, col, texto] para calcular altos
    }
    celda(f, c) { return this.ws.getCell(f, c); }
    combinar(f, c1, c2) { if (c2 > c1) this.ws.mergeCells(f, c1, f, c2); }
    titulo(t) {
      const c = this.celda(this.fila, 1);
      c.value = t; c.font = { bold: true, size: 15, color: { argb: "FFFFFFFF" } }; c.fill = rellenar("#0d366b");
      c.alignment = { vertical: "middle" };
      this.combinar(this.fila, 1, COLS_TEXTO); this.ws.getRow(this.fila).height = 28;
      this.fila += 2;
    }
    subtitulo(t) {
      const c = this.celda(this.fila, 1);
      c.value = t; c.font = { bold: true, size: 12, color: { argb: "FF1C5CAB" } };
      this.combinar(this.fila, 1, COLS_TEXTO);
      this.fila++;
    }
    texto(t, { relleno = null, negrita = false, desde = 1 } = {}) {
      if (!t) return;
      const c = this.celda(this.fila, desde);
      c.value = t; c.alignment = { wrapText: true, vertical: "top" };
      if (negrita) c.font = { bold: true };
      if (relleno) c.fill = rellenar(relleno);
      this.combinar(this.fila, desde, Math.max(desde, COLS_TEXTO));
      this.parrafos.push([this.fila, t, desde]);
      this.fila++;
    }
    par(etiqueta, contenido) {          // etiqueta en A, contenido en B…H
      const a = this.celda(this.fila, 1);
      a.value = etiqueta; a.font = { bold: true }; a.alignment = { wrapText: true, vertical: "top" };
      this.anchos[0] = Math.max(this.anchos[0], Math.min(34, etiqueta.length + 2));
      this.celdas.push([this.fila, 1, etiqueta]);
      if (contenido) this.texto(contenido, { desde: 2 }); else this.fila++;
    }
    filaTabla(valores, { cabecera = false, pie = false, colores = [] } = {}) {
      valores.forEach((txt, i) => {
        const c = this.celda(this.fila, i + 1);
        c.border = BORDES;
        c.alignment = { wrapText: true, vertical: "top" };
        if (colores[i]) { c.fill = rellenar(colores[i]); return; }
        if (cabecera) { c.value = txt; c.font = { bold: true }; c.fill = rellenar("#eef2f8"); }
        else {
          const { v, fmt } = valorCelda(txt);
          c.value = v; if (fmt) c.numFmt = fmt;
          if (pie) { c.font = { bold: true }; c.fill = rellenar("#f6f8fb"); }
        }
        this.anchos[i] = Math.max(this.anchos[i], Math.min(55, Math.max(...String(txt).split("\n").map(l => l.length)) + 2));
        this.celdas.push([this.fila, i + 1, String(txt)]);
      });
      this.fila++;
    }
    tabla(tablaEl) {
      const grupos = [["thead", { cabecera: true }], ["tbody", {}], ["tfoot", { pie: true }]];
      grupos.forEach(([sel, op]) => tablaEl.querySelectorAll(`${sel} tr`).forEach(tr => {
        const celdas = [...tr.children];
        const colores = celdas.map(td => { const p = td.querySelector(".punto-color"); return p ? cssAHex(p.style.background || getComputedStyle(p).backgroundColor) : null; });
        this.filaTabla(celdas.map(td => textoDe(td)), { ...op, colores });
      }));
      this.fila++;
    }
    imagen(url, ancho, alto) {
      const id = this.libro.addImage({ base64: url, extension: "png" });
      this.ws.addImage(id, { tl: { col: 0, row: this.fila - 1 }, ext: { width: ancho, height: alto } });
      this.fila += Math.ceil(alto / 20) + 1;
    }
    espacio() { this.fila++; }
    terminar() {
      const ws = this.ws;
      const anchos = this.anchos.map((a, i) => Math.max(i === 0 ? 16 : 11, a));
      anchos.forEach((a, i) => { ws.getColumn(i + 1).width = a; });
      const altos = {};
      const lineas = (t, ancho) => t.split("\n").reduce((s, l) => s + Math.max(1, Math.ceil(l.length / Math.max(1, ancho - 1))), 0);
      this.celdas.forEach(([f, c, t]) => { altos[f] = Math.max(altos[f] || 1, lineas(t, anchos[c - 1])); });
      this.parrafos.forEach(([f, t, desde]) => {
        const total = anchos.slice(desde - 1, COLS_TEXTO).reduce((a, b) => a + b, 0) * 1.1;
        altos[f] = Math.max(altos[f] || 1, lineas(t, total));
      });
      Object.entries(altos).forEach(([f, n]) => { if (n > 1) ws.getRow(Number(f)).height = Math.min(400, n * 15 + 3); });
    }
  }

  /** Dibuja un gráfico de la página en alta resolución con fondo blanco. */
  function imagenDeGrafico(g, ancho = 760, alto = 395) {
    const c = document.createElement("canvas");
    c.width = ancho; c.height = alto;
    const raw = g.config._config || g.config;
    const fondo = { id: "fondoBlanco", beforeDraw(ch) { const x = ch.ctx; x.save(); x.fillStyle = "#fff"; x.fillRect(0, 0, ch.width, ch.height); x.restore(); } };
    const tmp = new Chart(c, {
      type: raw.type,
      data: JSON.parse(JSON.stringify(raw.data)),
      options: { ...raw.options, responsive: false, maintainAspectRatio: false, animation: false, devicePixelRatio: 2 },
      plugins: [...(raw.plugins || []), fondo]
    });
    const url = tmp.toBase64Image("image/png", 1);
    tmp.destroy();
    return url;
  }
  /** Convierte el diagrama de caja (SVG) en imagen PNG. */
  function imagenDeSvg(svg, ancho = 1000) {
    return new Promise(resolver => {
      const vb = svg.viewBox.baseVal;
      const alto = Math.round(ancho * vb.height / vb.width);
      const img = new Image();
      img.onload = () => {
        const c = document.createElement("canvas"); c.width = ancho; c.height = alto;
        const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, ancho, alto); x.drawImage(img, 0, 0, ancho, alto);
        resolver({ url: c.toDataURL("image/png"), ancho, alto });
      };
      img.onerror = () => resolver(null);
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(svg));
    });
  }
  function graficoDeCanvas(canvas) { return Object.values(graficos).find(g => g.canvas === canvas); }

  /** Copia un elemento de la página (y su contenido) a la hoja. */
  async function copiarElemento(hoja, el) {
    if (!(el instanceof HTMLElement) && !(el instanceof SVGElement)) return;
    if (el.hidden || el.matches("button, script, .barra-exportar, .acciones")) return;
    const tag = el.tagName.toLowerCase();
    if (tag === "h2") { hoja.titulo(textoDe(el).replace(/^\d+\s*/, "")); return; }
    if (tag === "h3") { hoja.subtitulo(textoDe(el)); return; }
    if (tag === "table") { hoja.tabla(el); return; }
    if (tag === "canvas") {
      const g = graficoDeCanvas(el);
      if (g) hoja.imagen(imagenDeGrafico(g), 760, 395);
      return;
    }
    if (el.id === "gCaja") {
      const svg = el.querySelector("svg");
      const img = svg ? await imagenDeSvg(svg) : null;
      if (img) hoja.imagen(img.url, 760, Math.round(760 * img.alto / img.ancho));
      el.querySelectorAll("p").forEach(p => hoja.texto(textoDe(p)));
      return;
    }
    if (el.classList.contains("explicacion")) {
      const ps = el.querySelectorAll(":scope > p");
      if (ps.length) ps.forEach(p => hoja.texto(textoDe(p), { relleno: "#eaf2fd" }));
      else hoja.texto(textoDe(el), { relleno: "#eaf2fd" });
      hoja.espacio();
      return;
    }
    if (el.classList.contains("que-significa")) { hoja.texto(textoDe(el), { relleno: "#eaf2fd" }); hoja.espacio(); return; }
    if (el.classList.contains("formula-fila")) {
      const [a, b] = el.children;
      hoja.par(textoDe(a), textoDe(b));
      return;
    }
    if (el.classList.contains("tarjetas")) {
      hoja.filaTabla(["Medida", "Valor", "¿Qué significa?"], { cabecera: true });
      el.querySelectorAll(".tarjeta").forEach(t => hoja.filaTabla([textoDe(t.querySelector("span")), textoDe(t.querySelector("strong")), textoDe(t.querySelector("small"))]));
      hoja.espacio();
      return;
    }
    if (el.classList.contains("leyenda")) {
      hoja.subtitulo("Colores de los intervalos (los mismos en la tabla y en todos los gráficos)");
      el.querySelectorAll(".leyenda-item").forEach(it => {
        const p = it.querySelector(".punto-color");
        hoja.filaTabla(["", textoDe(it)], { colores: [cssAHex(p.style.background)] });
      });
      hoja.espacio();
      return;
    }
    if (el.classList.contains("formulas")) {
      hoja.filaTabla(["Concepto", "Fórmula", "Explicación"], { cabecera: true });
      el.querySelectorAll(".formula-item").forEach(it => hoja.filaTabla([textoDe(it.querySelector("strong")), textoDe(it.querySelector("code")), textoDe(it.querySelector("p"))]));
      return;
    }
    if (el.classList.contains("formas-ilustracion")) {
      const activa = el.querySelector(".forma-mini.activa");
      if (activa) hoja.texto("Forma de tus datos: " + textoDe(activa).replace("← tus datos", "").trim(), { negrita: true });
      return;
    }
    if (el.classList.contains("ordenados")) {
      hoja.texto("La lista completa de datos ordenados (con su posición) está en la hoja «Datos».");
      return;
    }
    if (tag === "label") {
      const control = el.querySelector("textarea, input:not([type=checkbox]):not([type=file]), select");
      const check = el.querySelector("input[type=checkbox]");
      const propio = [...el.childNodes].filter(n => n.nodeType === 3 || (n.nodeType === 1 && !n.matches("textarea, input, select")))
        .map(n => n.textContent).join(" ").replace(/\s+/g, " ").trim();
      if (check) { hoja.texto(`${propio}: ${check.checked ? "sí" : "no"}`); return; }
      if (control) {
        const valor = control.tagName === "SELECT" ? control.options[control.selectedIndex].text : control.value.trim().split(/\r?\n/).join("  |  ");
        hoja.par(propio, valor);
      }
      return;
    }
    if (tag === "details") {
      const s = el.querySelector("summary");
      hoja.subtitulo(s ? textoDe(s) : "Explicación");
      for (const hijo of el.children) if (hijo !== s) await copiarElemento(hoja, hijo);
      hoja.espacio();
      return;
    }
    if (tag === "dl") {
      el.querySelectorAll("dt").forEach(dt => { const dd = dt.nextElementSibling; hoja.par(textoDe(dt), dd ? textoDe(dd) : ""); });
      return;
    }
    if (tag === "ol" || tag === "ul") {
      [...el.children].forEach((li, i) => hoja.texto(`${tag === "ol" ? (i + 1) + "." : "•"} ${textoDe(li)}`));
      return;
    }
    if (tag === "p" || tag === "figcaption") { hoja.texto(textoDe(el)); if (tag === "figcaption") hoja.espacio(); return; }
    // Contenedores (section, div, figure…): copiar sus hijos en orden
    for (const hijo of el.children) await copiarElemento(hoja, hijo);
  }

  const HOJAS = [
    ["resumen", "Resumen"], ["pasos", "Paso a paso"], ["tabla", "Tabla de frecuencias"], ["graficos", "Gráficos"],
    ["central", "Tendencia central"], ["posicion", "Posición"], ["dispersion", "Dispersión"], ["forma", "Forma"],
    ["regresion", "Regresión"], ["promedios", "Los cuatro promedios"], ["glosario", "Glosario"]
  ];

  /** Arma el libro de Excel completo (se separa de la descarga para poder probarlo). */
  async function construirLibro() {
    const t = titulo();
    const libro = new ExcelJS.Workbook();
    libro.creator = "Calculadora de distribución de frecuencias";
    libro.created = new Date();

    // Portada
    const portada = new Hoja(libro, "Portada");
    portada.titulo("Calculadora de distribución de frecuencias");
    portada.par("Variable analizada", t);
    portada.par("Cantidad de datos (n)", String(R.n));
    portada.par("Fecha", new Date().toLocaleString("es-CO"));
    portada.par("Decimales mostrados", String(dec()));
    portada.espacio();
    portada.subtitulo("Contenido del archivo");
    ["Datos", ...HOJAS.map(h => h[1])].forEach((n, i) => portada.par(`Hoja ${i + 2}`, n));
    portada.espacio();
    portada.texto("Cada hoja tiene la fórmula, los números reemplazados, el resultado y una explicación en palabras sencillas. Cada intervalo conserva el mismo color en la tabla y en los gráficos.", { relleno: "#eaf2fd" });
    portada.terminar();

    // Datos (original y ordenado)
    const hd = new Hoja(libro, "Datos");
    hd.titulo(`Datos: ${t}`);
    hd.filaTabla(["N° (orden escrito)", t, "Posición (ordenados)", "Dato ordenado"], { cabecera: true });
    const fmt = "#,##0" + (R.dec ? "." + "0".repeat(R.dec) : "");
    R.valores.forEach((v, i) => {
      const fila = hd.fila;
      [i + 1, v, i + 1, R.ordenados[i]].forEach((val, c) => {
        const cel = hd.celda(fila, c + 1);
        cel.value = val; cel.border = BORDES;
        if (c % 2 === 1) cel.numFmt = fmt;
      });
      hd.fila++;
    });
    hd.anchos = [20, Math.max(16, t.length + 2), 22, 16];
    hd.celdas = hd.celdas.filter(([f]) => f !== 3);
    hd.terminar();

    // Una hoja por cada sección de la página
    for (const [id, nombre] of HOJAS) {
      const hoja = new Hoja(libro, nombre);
      await copiarElemento(hoja, $(id));
      hoja.terminar();
    }
    return libro;
  }

  async function descargarExcel() {
    if (!R) return;
    if (!window.ExcelJS) { alert("No se pudo cargar la librería de Excel. Revisa tu conexión a internet y vuelve a intentarlo."); return; }
    const boton = $("btnExcel");
    const textoBoton = boton.textContent;
    boton.disabled = true; boton.textContent = "Generando Excel…";
    try {
      const libro = await construirLibro();
      const buffer = await libro.xlsx.writeBuffer();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      a.download = `distribucion_${titulo().replace(/[^\wáéíóúñ]+/gi, "_").slice(0, 40)}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    } catch (err) {
      console.error(err);
      alert("Ocurrió un error al crear el Excel: " + err.message);
    } finally {
      boton.disabled = false; boton.textContent = textoBoton;
    }
  }
  window.__construirLibro = construirLibro;   // usado solo para pruebas

  /* ---------------------------- Eventos ------------------------------- */
  function cargarEjemplo(clave) {
    const ej = EJEMPLOS[clave];
    $("titulo").value = ej.titulo; $("titulo").dataset.auto = "0";
    $("entrada").value = ej.datos.join("\n");
    actualizarConteo();
    calcularTodo();
  }

  $("entrada").addEventListener("input", actualizarConteo);
  $("btnCalcular").addEventListener("click", () => calcularTodo());
  $("btnEjemplo1").addEventListener("click", () => cargarEjemplo("unidades"));
  $("btnEjemplo2").addEventListener("click", () => cargarEjemplo("precio"));
  $("btnLimpiar").addEventListener("click", () => {
    $("entrada").value = ""; R = null; destruir();
    $("resultados").hidden = true; actualizarConteo();
  });
  $("archivo").addEventListener("change", e => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (file.size > 5 * 1024 * 1024) { alert("El archivo pesa más de 5 MB."); return; }
    const lector = new FileReader();
    lector.onload = () => { $("titulo").dataset.auto = "1"; $("entrada").value = String(lector.result); actualizarConteo(); calcularTodo(); };
    lector.readAsText(file);
    e.target.value = "";
  });
  $("titulo").addEventListener("input", () => { $("titulo").dataset.auto = "0"; if (R) calcularTodo(false); });
  $("decimales").addEventListener("change", () => { if (R) calcularTodo(false); });
  $("btnPk").addEventListener("click", calcPk);
  $("pk").addEventListener("keydown", e => { if (e.key === "Enter") calcPk(); });
  $("btnExcel").addEventListener("click", descargarExcel);
  $("btnPdf").addEventListener("click", () => window.print());
  window.addEventListener("beforeprint", () => Object.values(graficos).forEach(g => g.resize()));

  actualizarConteo();
})();
