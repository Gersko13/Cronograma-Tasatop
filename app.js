// app.js
(() => {
  "use strict";

  // =========================
  // CONSTANTES (VBA)
  // =========================
  const TASA_IMPUESTO_2DA = 0.05;
 const LOGO_URL = "https://images.weserv.nl/?url=tasatop.com.pe/wp-content/uploads/elementor/thumbs/logos-17-r320c27cra7m7te2fafiia4mrbqd3aqj7ifttvy33g.png";
;

  // =========================
  // UTILIDADES FECHA (UTC para evitar DST)
  // =========================
  const MS_PER_DAY = 24 * 60 * 60 * 1000;

  function makeUTCDate(y, m1, d) {
    // m1 = 1..12
    return new Date(Date.UTC(y, m1 - 1, d, 0, 0, 0, 0));
  }

  function toUTCDateFromInput(value) {
    // value: "YYYY-MM-DD"
    if (!value || typeof value !== "string") return null;
    const parts = value.split("-");
    if (parts.length !== 3) return null;
    const y = Number(parts[0]);
    const m = Number(parts[1]);
    const d = Number(parts[2]);
    if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
    const dt = makeUTCDate(y, m, d);
    // Validación estricta (evita overflow tipo 2026-02-31)
    if (dt.getUTCFullYear() !== y || (dt.getUTCMonth() + 1) !== m || dt.getUTCDate() !== d) return null;
    return dt;
  }

  function dateDiffDaysUTC(d1, d2) {
    // Replica DateDiff("d", d1, d2) para fechas a medianoche
    const t1 = Date.UTC(d1.getUTCFullYear(), d1.getUTCMonth(), d1.getUTCDate());
    const t2 = Date.UTC(d2.getUTCFullYear(), d2.getUTCMonth(), d2.getUTCDate());
    return Math.round((t2 - t1) / MS_PER_DAY);
  }

  function addMonthsUTC(dateUTC, months) {
    const y = dateUTC.getUTCFullYear();
    const m0 = dateUTC.getUTCMonth(); // 0..11
    const d = dateUTC.getUTCDate();

    const targetMonthIndex = m0 + months;
    const y2 = y + Math.floor(targetMonthIndex / 12);
    const m2 = ((targetMonthIndex % 12) + 12) % 12;

    // En VBA, FechaPagoMes usa DateSerial(y,m, dAjustado) con último día si no existe.
    // Aquí esta función solo ajusta mes manteniendo día si posible; si no, cae al último día.
    const lastDay = new Date(Date.UTC(y2, m2 + 1, 0)).getUTCDate();
    const d2 = Math.min(d, lastDay);
    return new Date(Date.UTC(y2, m2, d2));
  }

  function formatDateDDMMYYYY(dateUTC) {
    if (!dateUTC) return "";
    const dd = String(dateUTC.getUTCDate()).padStart(2, "0");
    const mm = String(dateUTC.getUTCMonth() + 1).padStart(2, "0");
    const yyyy = dateUTC.getUTCFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  function nowStampYYYYMMDD_HHMMSS() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const HH = String(d.getHours()).padStart(2, "0");
    const MM = String(d.getMinutes()).padStart(2, "0");
    const SS = String(d.getSeconds()).padStart(2, "0");
    return `${yyyy}${mm}${dd}_${HH}${MM}${SS}`;
  }

  function formatGeneratedAt() {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    const HH = String(d.getHours()).padStart(2, "0");
    const MM = String(d.getMinutes()).padStart(2, "0");
    const SS = String(d.getSeconds()).padStart(2, "0");
    return `${dd}/${mm}/${yyyy} ${HH}:${MM}:${SS}`;
  }

  // =========================
  // ROUND VBA (Banker's rounding)
  // VBA Round(x,2) usa "round half to even".
  // =========================
  function roundBankers(value, digits) {
    if (!Number.isFinite(value)) return 0;
    const factor = Math.pow(10, digits);
    const n = value * factor;

    // Evitar errores binarios: trabaja con epsilon relativa
    const eps = 1e-12 * Math.max(1, Math.abs(n));
    const sign = n < 0 ? -1 : 1;
    const abs = Math.abs(n);

    const floor = Math.floor(abs + eps);
    const frac = abs - floor;

    if (Math.abs(frac - 0.5) <= 1e-10) {
      // Half: to even
      const isEven = (floor % 2) === 0;
      const rounded = isEven ? floor : floor + 1;
      return sign * (rounded / factor);
    }
    return sign * (Math.round(abs + eps) / factor);
  }

  // =========================
  // NORMALIZACIÓN (VBA NormalizarClave)
  // =========================
  function normalizarClave(s) {
    s = String(s ?? "").trim().toUpperCase();

    // Sin tildes / Ñ
    const map = {
      "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
      "Ü": "U", "Ñ": "N"
    };
    s = s.replace(/[ÁÉÍÓÚÜÑ]/g, ch => map[ch] || ch);

    // Espacios dobles (en VBA la intención es colapsar)
    s = s.replace(/\s+/g, " ").trim();
    return s;
  }

  // =========================
  // REGLAS DE NEGOCIO (VBA)
  // =========================
  function obtenerDiaPago(producto) {
    switch (normalizarClave(producto)) {
      case "IKB": return 15;
      case "ALI": return 28;
      case "PET": return 10;
      case "M&L": return 20;
      default: return 15;
    }
  }

  function frecuenciaAMeses(frecuencia, plazoMeses) {
    switch (normalizarClave(frecuencia)) {
      case "MENSUAL": return 1;
      case "BIMESTRAL": return 2;
      case "TRIMESTRAL": return 3;
      case "SEMESTRAL": return 6;
      case "ANUAL": return 12;
      case "AL FINALIZAR": return plazoMeses;
      default: return 1;
    }
  }

  function fechaPagoMes(fechaBaseUTC, mesOffset, diaPago) {
    // VBA:
    // y = Year(DateAdd("m", mesOffset, fechaBase))
    // m = Month(DateAdd("m", mesOffset, fechaBase))
    // ultimoDia = Day(DateSerial(y, m+1, 0))
    // d = min(diaPago, ultimoDia)
    const shifted = addMonthsUTC(fechaBaseUTC, mesOffset);
    const y = shifted.getUTCFullYear();
    const m = shifted.getUTCMonth() + 1; // 1..12
    const ultimoDia = new Date(Date.UTC(y, m, 0)).getUTCDate(); // Date.UTC(y, m, 0) => último día del mes m
    const d = Math.min(diaPago, ultimoDia);
    return makeUTCDate(y, m, d);
  }

  function calcularPrimeraFechaPago(fechaInicioUTC, diaPago, opcionPrimerPago) {
    const op = normalizarClave(opcionPrimerPago);

    if (fechaInicioUTC.getUTCDate() > diaPago) {
      return fechaPagoMes(fechaInicioUTC, 1, diaPago);
    }

    // "MES" y "INVERSION" en cualquier variación
    if (op.includes("MES") && op.includes("INVERSION")) {
      return fechaPagoMes(fechaInicioUTC, 0, diaPago);
    }
    return fechaPagoMes(fechaInicioUTC, 1, diaPago);
  }

  function esMesDePago(i, freqMeses) {
    if (freqMeses <= 0) return false;
    return (i % freqMeses) === 0;
  }

  // =========================
  // FORMATO NUMÉRICO
  // =========================
  function formatNumber2(n) {
    const v = Number(n);
    if (!Number.isFinite(v)) return "0.00";
    // En UI: #,##0.00
    return v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function money(moneda, n) {
    return `${moneda} ${formatNumber2(n)}`;
  }

  function formatTA(tasaEA) {
    // VBA: Format(tasaEA,"0.000%")
    // tasaEA ya viene como decimal (ej 0.12)
    const pct = tasaEA * 100;
    // 3 decimales en porcentaje
    const s = pct.toLocaleString("en-US", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
    return `${s}%`;
  }

  // =========================
  // LÓGICA PRINCIPAL (replica 1:1 VBA GenerarCronograma)
  // =========================
  function generarCronograma(inputs) {
    const {
      fechaInicioUTC,
      monto,
      moneda,
      tasaEA,        // decimal (ej 0.12)
      plazo,         // meses (int)
      producto,
      frecuenciaInteres,
      frecuenciaCapital,
      opcionPrimerPago
    } = inputs;

    const diaPago = obtenerDiaPago(producto);
    const freqIntMeses = frecuenciaAMeses(frecuenciaInteres, plazo);
    const freqCapMeses = frecuenciaAMeses(frecuenciaCapital, plazo);

    // >>> B10 reflejado en Mes 1:
    let fechaPagoAnterior = calcularPrimeraFechaPago(fechaInicioUTC, diaPago, opcionPrimerPago);

    // Inicialización
    let saldo = monto;
    let pagosInteresCont = 0;
    let pagosCapitalCont = 0;

    let numPagosCapital;
    if (normalizarClave(frecuenciaCapital) === "AL FINALIZAR") {
      numPagosCapital = 1;
    } else {
      numPagosCapital = Math.floor((plazo + freqCapMeses - 1) / freqCapMeses); // ceiling
      if (numPagosCapital < 1) numPagosCapital = 1;
    }
    const amortBase = monto / numPagosCapital;

    const rows = [];

    for (let i = 0; i <= plazo; i++) {
      const mes = i;

      let fechaPagoUTC = null;
      let fechaCronoUTC = null;

      let diasInfo = 0;
      let diasInteres = 0;

      let pagaInteres = false;
      let pagaCapital = false;

      let interesMes = 0;
      let interesBrutoPago = 0;
      let impuesto = 0;
      let interesDepositar = 0;

      let devolucionCapital = 0;

      if (i === 0) {
        // Mes 0: sin pago
        fechaPagoUTC = null;
        fechaCronoUTC = fechaInicioUTC;
        diasInfo = 0;
      } else {
        // Mes 1..plazo
        if (i === 1) {
          fechaPagoUTC = fechaPagoAnterior;
        } else {
          fechaPagoUTC = fechaPagoMes(fechaPagoAnterior, 1, diaPago);
          fechaPagoAnterior = fechaPagoUTC;
        }

        // CLAVE: Cronograma = Pago siempre (desde mes 1)
        fechaCronoUTC = fechaPagoUTC;

        // Días informativos
        if (i === 1) {
          diasInfo = dateDiffDaysUTC(fechaInicioUTC, fechaPagoUTC);
          if (diasInfo < 0) diasInfo = 0;
        } else {
          const fechaPagoPrev = fechaPagoMes(fechaPagoUTC, -1, diaPago);
          diasInfo = dateDiffDaysUTC(fechaPagoPrev, fechaPagoUTC);
          if (diasInfo <= 0) diasInfo = 30;
        }
      }

      // INTERÉS
      if (i === 0) {
        pagaInteres = false;
      } else {
        if (freqIntMeses === 1) {
          pagaInteres = true;
        } else if (esMesDePago(mes, freqIntMeses) || mes === plazo) {
          pagaInteres = true;
        } else {
          pagaInteres = false;
        }
      }

      if (pagaInteres) {
        if (pagosInteresCont === 0) {
          diasInteres = dateDiffDaysUTC(fechaInicioUTC, fechaPagoUTC);
          if (diasInteres < 0) diasInteres = 0;
        } else {
          diasInteres = 30 * freqIntMeses;
        }

        interesBrutoPago = (((1 + tasaEA) ** (diasInteres / 360)) - 1) * saldo;
        interesBrutoPago = roundBankers(interesBrutoPago, 2);
        impuesto = roundBankers(interesBrutoPago * TASA_IMPUESTO_2DA, 2);
        interesDepositar = roundBankers(interesBrutoPago - impuesto, 2);
        interesMes = interesBrutoPago;

        pagosInteresCont++;
      } else {
        diasInteres = 0;
        interesMes = 0;
        interesBrutoPago = 0;
        impuesto = 0;
        interesDepositar = 0;
      }

      // CAPITAL
      if (i === 0) {
        pagaCapital = false;
      } else {
        if (normalizarClave(frecuenciaCapital) === "AL FINALIZAR") {
          pagaCapital = (mes === plazo);
        } else {
          pagaCapital = esMesDePago(mes, freqCapMeses) || (mes === plazo);
        }
      }

      // Monto base = saldo al inicio del periodo
      const saldoInicio = saldo;

      if (pagaCapital && saldo > 0) {
        pagosCapitalCont++;

        if (mes === plazo || pagosCapitalCont === numPagosCapital) {
          devolucionCapital = saldo;
        } else {
          devolucionCapital = roundBankers(amortBase, 2);
          if (devolucionCapital > saldo) devolucionCapital = saldo;
        }

        saldo = roundBankers(saldo - devolucionCapital, 2);
        if (saldo < 0) saldo = 0;
      } else {
        devolucionCapital = 0;
      }

      const totalDepositar = roundBankers(interesDepositar + devolucionCapital, 2);

      rows.push({
        mes: i,
        fechaCronoUTC: fechaCronoUTC,
        fechaPagoUTC: i === 0 ? null : fechaPagoUTC,
        dias: i === 0 ? "" : (pagaInteres ? diasInteres : diasInfo),

        montoBase: roundBankers(saldoInicio, 2),
        interesBruto: interesMes,
        impuesto: impuesto,
        interesDepositar: interesDepositar,
        devolucionCapital: devolucionCapital,
        saldoCapital: saldo,
        totalDepositar: totalDepositar,

        // extras para PDF/UI si se requieren
        _meta: {
          pagaInteres,
          pagaCapital,
          diasInteres,
          diasInfo,
          diaPago,
          freqIntMeses,
          freqCapMeses
        }
      });
    }

    // Totales (VBA suma columnas H, I, K)
    const totalInteresDepositar = roundBankers(rows.reduce((acc, r) => acc + (Number(r.interesDepositar) || 0), 0), 2);
    const totalDevCap = roundBankers(rows.reduce((acc, r) => acc + (Number(r.devolucionCapital) || 0), 0), 2);
    const totalDepositar = roundBankers(rows.reduce((acc, r) => acc + (Number(r.totalDepositar) || 0), 0), 2);

    return {
      rows,
      totals: {
        interesDepositar: totalInteresDepositar,
        devolucionCapital: totalDevCap,
        totalDepositar: totalDepositar
      },
      meta: {
        diaPago,
        freqIntMeses,
        freqCapMeses,
        numPagosCapital,
        amortBase
      }
    };
  }

  // =========================
  // VALIDACIONES UI (equivalentes)
  // =========================
  function clearErrors() {
    const errs = document.querySelectorAll(".field__error");
    errs.forEach(e => (e.textContent = ""));
    const fields = document.querySelectorAll("input,select");
    fields.forEach(f => f.classList.remove("is-invalid"));
    $("#alerts").innerHTML = "";
  }

  function setFieldError(id, msg) {
    const el = document.getElementById(id);
    if (el) el.classList.add("is-invalid");
    const err = document.getElementById("err_" + id);
    if (err) err.textContent = msg;
  }

  function showTopError(title, text) {
    const alerts = $("#alerts");
    const div = document.createElement("div");
    div.className = "alert";
    div.innerHTML = `
      <span class="alert__dot"></span>
      <div>
        <p class="alert__title">${escapeHtml(title)}</p>
        <p class="alert__text">${escapeHtml(text)}</p>
      </div>
    `;
    alerts.appendChild(div);
  }

  function validateInputs(raw) {
    const errors = [];

    // B2
    if (!raw.fechaInicioUTC) {
      errors.push({ field: "fechaInicio", msg: "La fecha de inicio no es una fecha válida." });
    }

    // B3
    if (!Number.isFinite(raw.monto) || raw.monto <= 0) {
      errors.push({ field: "monto", msg: "El monto debe ser numérico y mayor a 0." });
    }

    // B4
    if (!raw.moneda || (raw.moneda !== "S/" && raw.moneda !== "$")) {
      errors.push({ field: "moneda", msg: "La moneda está vacía o no es válida. Usa S/ o $." });
    }

    // B5
    if (!Number.isFinite(raw.tasaEA_pct)) {
      errors.push({ field: "tasaEA", msg: "La tasa debe ser numérica." });
    }

    // B6
    if (!Number.isFinite(raw.plazo) || raw.plazo <= 0 || !Number.isInteger(raw.plazo)) {
      errors.push({ field: "plazo", msg: "El plazo en meses debe ser numérico entero y mayor a 0." });
    }

    // B7
    if (!raw.producto) {
      errors.push({ field: "producto", msg: "El producto está vacío." });
    }

    // B8
    if (!raw.frecuenciaInteres) {
      errors.push({ field: "freqInteres", msg: "La frecuencia de intereses está vacía." });
    }

    // B9
    if (!raw.frecuenciaCapital) {
      errors.push({ field: "freqCapital", msg: "La devolución de capital está vacía." });
    }

    // B10 (si vacío => "Próximo mes")
    return errors;
  }

  // =========================
  // RENDER UI
  // =========================
  function renderSummary(inputs) {
    const el = $("#summary");
    const html = `
      <div class="kv">
        <div class="kv__item">
          <div class="kv__label">TA</div>
          <div class="kv__value kv__value--brand">${escapeHtml(formatTA(inputs.tasaEA))}</div>
        </div>
        <div class="kv__item">
          <div class="kv__label">Monto invertido</div>
          <div class="kv__value">${escapeHtml(money(inputs.moneda, inputs.monto))}</div>
        </div>
        <div class="kv__item">
          <div class="kv__label">Producto</div>
          <div class="kv__value">${escapeHtml(inputs.producto)}</div>
        </div>
        <div class="kv__item">
          <div class="kv__label">Plazo</div>
          <div class="kv__value">${escapeHtml(String(inputs.plazo))} Meses</div>
        </div>
        <div class="kv__item">
          <div class="kv__label">Frecuencia (interés)</div>
          <div class="kv__value">${escapeHtml(inputs.frecuenciaInteres)}</div>
        </div>
        <div class="kv__item">
          <div class="kv__label">Tipo tasa</div>
          <div class="kv__value">Tasa Efectiva Anual</div>
        </div>
        <div class="kv__item" style="grid-column:1 / -1;">
          <div class="kv__label">Devolución de capital</div>
          <div class="kv__value">${escapeHtml(inputs.frecuenciaCapital)}</div>
        </div>
      </div>
    `;
    el.innerHTML = html;
  }

  function renderTable(result, moneda) {
    const tbody = $("#tbody");
    const tfoot = $("#tfoot");
    tbody.innerHTML = "";
    tfoot.innerHTML = "";

    result.rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.mes}</td>
        <td>${escapeHtml(formatDateDDMMYYYY(r.fechaCronoUTC))}</td>
        <td>${escapeHtml(formatDateDDMMYYYY(r.fechaPagoUTC))}</td>
        <td>${escapeHtml(String(r.dias))}</td>
        <td>${escapeHtml(formatNumber2(r.montoBase))}</td>
        <td>${escapeHtml(formatNumber2(r.interesBruto))}</td>
        <td>${escapeHtml(formatNumber2(r.impuesto))}</td>
        <td>${escapeHtml(formatNumber2(r.interesDepositar))}</td>
        <td>${escapeHtml(formatNumber2(r.devolucionCapital))}</td>
        <td>${escapeHtml(formatNumber2(r.saldoCapital))}</td>
        <td>${escapeHtml(formatNumber2(r.totalDepositar))}</td>
      `;
      tbody.appendChild(tr);
    });

    // Totales (como VBA: TOTAL en A, sumas en H, I, K)
    const trf = document.createElement("tr");
    trf.innerHTML = `
      <td>TOTAL</td>
      <td colspan="6"></td>
      <td>${escapeHtml(formatNumber2(result.totals.interesDepositar))}</td>
      <td>${escapeHtml(formatNumber2(result.totals.devolucionCapital))}</td>
      <td></td>
      <td>${escapeHtml(formatNumber2(result.totals.totalDepositar))}</td>
    `;
    tfoot.appendChild(trf);

    $("#rowCount").textContent = `${result.rows.length} filas (incluye Mes 0)`;
  }

  // =========================
  // PDF (A4 horizontal, autoscale, membrete con logo embebido)
  // =========================
  let cachedLogoDataUrl = null;

  async function getLogoDataUrl() {
    if (cachedLogoDataUrl) return cachedLogoDataUrl;

    // Intento 1: fetch -> blob -> FileReader (requiere CORS permitido por el servidor)
    try {
      const res = await fetch(LOGO_URL, { mode: "cors", cache: "force-cache" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const dataUrl = await blobToDataURL(blob);
      cachedLogoDataUrl = dataUrl;
      return dataUrl;
    } catch (e) {
      // Intento 2: usar <img> ya cargada y canvas (también requiere CORS OK)
      try {
        const img = document.getElementById("brandLogo");
        const dataUrl = await imageElToDataURL(img);
        cachedLogoDataUrl = dataUrl;
        return dataUrl;
      } catch (e2) {
        // Si el servidor bloquea CORS, no se puede convertir desde el navegador.
        // La exportación continúa con membrete en texto (sin “logo roto”).
        return null;
      }
    }
  }

  function blobToDataURL(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  function imageElToDataURL(img) {
    return new Promise((resolve, reject) => {
      try {
        if (!img || !img.complete) throw new Error("Imagen no cargada");
        const canvas = document.createElement("canvas");
        canvas.width = img.naturalWidth || 1;
        canvas.height = img.naturalHeight || 1;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      } catch (err) {
        reject(err);
      }
    });
  }

  async function exportPDF(state) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

    const pageW = doc.internal.pageSize.getWidth();
    const margin = 28;

    // Membrete (UI + PDF)
    const gen = formatGeneratedAt();
    const titleY = 34;

    // Logo embebido (si CORS permite)
    const logoDataUrl = await getLogoDataUrl();
    let logoW = 0;
    let logoH = 0;

    if (logoDataUrl) {
      // Tamaño fijo visual
      logoW = 62;
      logoH = 62;
      doc.addImage(logoDataUrl, "PNG", margin, 16, logoW, logoH);
    }

    const textX = margin + (logoDataUrl ? (logoW + 12) : 0);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("TASATOP", textX, titleY);

    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    doc.text("Cronograma de Inversión", textX, titleY + 18);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10.5);
    doc.text(`Generado: ${gen}`, textX, titleY + 34);

    // Resumen (tabla compacta)
    const sum = state.inputs;
    const resumenBody = [
      ["TA", formatTA(sum.tasaEA), "Monto invertido", money(sum.moneda, sum.monto), "Producto", sum.producto],
      ["Plazo", `${sum.plazo} Meses`, "Frecuencia (interés)", sum.frecuenciaInteres, "Tipo tasa", "Tasa Efectiva Anual"],
      ["Devolución de capital", sum.frecuenciaCapital, "Opción primer pago (B10)", sum.opcionPrimerPago, "Día de pago", String(obtenerDiaPago(sum.producto))]
    ];

    doc.autoTable({
      startY: 92,
      margin: { left: margin, right: margin },
      theme: "grid",
      styles: { font: "helvetica", fontSize: 9.5, cellPadding: 4, halign: "left" },
      headStyles: { fillColor: [31, 95, 168], textColor: 255, fontStyle: "bold" },
      bodyStyles: { textColor: 20 },
      tableLineWidth: 0.5,
      tableLineColor: [220, 224, 231],
      head: [["Campo", "Valor", "Campo", "Valor", "Campo", "Valor"]],
      body: resumenBody
    });

    const startY = doc.lastAutoTable.finalY + 12;

    // Cronograma
    const head = [[
      "Mes", "Fecha cronograma", "Fecha pago", "Días",
      "Monto base", "Interés bruto", "Impuesto", "Interés a depositar",
      "Devolución capital", "Saldo capital", "Total a depositar"
    ]];

    const body = state.result.rows.map(r => ([
      String(r.mes),
      formatDateDDMMYYYY(r.fechaCronoUTC),
      r.mes === 0 ? "" : formatDateDDMMYYYY(r.fechaPagoUTC),
      r.mes === 0 ? "" : String(r.dias),
      formatNumber2(r.montoBase),
      formatNumber2(r.interesBruto),
      formatNumber2(r.impuesto),
      formatNumber2(r.interesDepositar),
      formatNumber2(r.devolucionCapital),
      formatNumber2(r.saldoCapital),
      formatNumber2(r.totalDepositar)
    ]));

    // Totales al final (igual que VBA: H, I, K)
    body.push([
      "TOTAL", "", "", "", "", "", "",
      formatNumber2(state.result.totals.interesDepositar),
      formatNumber2(state.result.totals.devolucionCapital),
      "",
      formatNumber2(state.result.totals.totalDepositar)
    ]);

    doc.autoTable({
      startY,
      margin: { left: margin, right: margin },
      theme: "grid",
      head,
      body,
      styles: { font: "helvetica", fontSize: 8.2, cellPadding: 3, halign: "center", valign: "middle" },
      headStyles: { fillColor: [31, 95, 168], textColor: 255, fontStyle: "bold" },
      alternateRowStyles: { fillColor: [245, 247, 251] },
      tableLineWidth: 0.5,
      tableLineColor: [220, 224, 231],
      didParseCell: function (data) {
        // Totales (última fila)
        const isLast = data.row.index === body.length - 1;
        if (isLast) {
          data.cell.styles.fontStyle = "bold";
          data.cell.styles.fillColor = [235, 238, 245];
        }
      }
    });

    // FitToPagesWide=1 (VBA) -> aquí garantizamos ancho con landscape + márgenes + autoTable compacto.
    const filename = `TASATOP_Cronograma_${nowStampYYYYMMDD_HHMMSS()}.pdf`;
    doc.save(filename);
  }

  // =========================
  // ESTADO + EVENTOS
  // =========================
  const state = {
    inputs: null,
    result: null
  };

  function readForm() {
    const fechaInicioUTC = toUTCDateFromInput($("#fechaInicio").value);

    const monto = Number($("#monto").value);
    const moneda = ($("#moneda").value || "").trim();

    const tasaEA_pct = Number($("#tasaEA").value);
    const tasaEA = tasaEA_pct / 100;

    const plazo = Number($("#plazo").value);

    let producto = ($("#producto").value || "").trim();
    if (producto === "OTRO") producto = "OTRO";

    const frecuenciaInteres = ($("#freqInteres").value || "").trim();
    const frecuenciaCapital = ($("#freqCapital").value || "").trim();

    let opcionPrimerPago = ($("#primerPago").value || "").trim();
    if (!opcionPrimerPago) opcionPrimerPago = "Próximo mes";

    return {
      fechaInicioUTC,
      monto,
      moneda,
      tasaEA_pct,
      tasaEA,
      plazo,
      producto,
      frecuenciaInteres,
      frecuenciaCapital,
      opcionPrimerPago
    };
  }

  function onGenerate() {
    clearErrors();
    const raw = readForm();
    const errs = validateInputs(raw);

    if (errs.length) {
      showTopError("Revisa los datos", "Corrige los campos marcados para generar el cronograma.");
      errs.forEach(e => setFieldError(e.field, e.msg));
      $("#btnPdf").disabled = true;
      return;
    }

    // Generar (replica VBA)
    const result = generarCronograma({
      fechaInicioUTC: raw.fechaInicioUTC,
      monto: raw.monto,
      moneda: raw.moneda,
      tasaEA: raw.tasaEA,
      plazo: raw.plazo,
      producto: raw.producto,
      frecuenciaInteres: raw.frecuenciaInteres,
      frecuenciaCapital: raw.frecuenciaCapital,
      opcionPrimerPago: raw.opcionPrimerPago
    });

    state.inputs = {
      ...raw,
      monto: raw.monto,
      plazo: raw.plazo
    };
    state.result = result;

    renderSummary({
      tasaEA: raw.tasaEA,
      moneda: raw.moneda,
      monto: raw.monto,
      producto: raw.producto,
      plazo: raw.plazo,
      frecuenciaInteres: raw.frecuenciaInteres,
      frecuenciaCapital: raw.frecuenciaCapital,
      opcionPrimerPago: raw.opcionPrimerPago
    });

    renderTable(result, raw.moneda);

    // Aviso si el logo no se puede convertir (CORS)
    // (No bloquea generación; PDF saldrá sin “logo roto”, solo sin imagen)
    getLogoDataUrl().then(d => {
      if (!d) {
        // Mensaje informativo (sin romper)
        showTopError(
          "Logo en PDF",
          "El servidor del logo no permite conversión a imagen embebida (CORS). El PDF se exportará con membrete en texto. Si TASATOP habilita CORS, el logo saldrá embebido automáticamente."
        );
      }
    });

    $("#btnPdf").disabled = false;
  }

  function onClear() {
    clearErrors();
    $("#form").reset();
    state.inputs = null;
    state.result = null;

    $("#summary").innerHTML = `<div class="summary__empty">Genera el cronograma para ver el resumen.</div>`;
    $("#tbody").innerHTML = `<tr><td colspan="11" class="td-empty">Aún no hay datos. Completa el formulario y genera el cronograma.</td></tr>`;
    $("#tfoot").innerHTML = "";
    $("#rowCount").textContent = "—";
    $("#btnPdf").disabled = true;
  }

  async function onPdf() {
    if (!state.result || !state.inputs) return;
    await exportPDF(state);
  }

  // =========================
  // HELPERS DOM
  // =========================
  function $(sel) {
    return document.querySelector(sel);
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  // =========================
  // INIT
  // =========================
  function init() {
    $("#generatedAt").textContent = `Generado: ${formatGeneratedAt()}`;

    // Precarga/intent de logo a dataURL para PDF
    // (si CORS permite, quedará cacheado)
    getLogoDataUrl().catch(() => { /* no-op */ });

    $("#btnGenerate").addEventListener("click", onGenerate);
    $("#btnClear").addEventListener("click", onClear);
    $("#btnPdf").addEventListener("click", onPdf);

    // UX: Enter => generar
    $("#form").addEventListener("submit", (e) => {
      e.preventDefault();
      onGenerate();
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();
