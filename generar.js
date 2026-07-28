/**
 * generar.js
 * Motor de renderizado PptxGenJS para la presentación Traçado + Perú 2026-2035
 * 
 * Requisitos previos:
 *   npm init -y
 *   npm install pptxgenjs
 * 
 * Ejecución:
 *   node generar.js
 * 
 * Salida:
 *   Traçado_Peru_2026-2035.pptx
 */

const PptxGenJS = require("pptxgenjs");
const datos = require("./datos.json");

const pptx = new PptxGenJS();

// ============================================================
// 1. CONFIGURACIÓN GLOBAL
// ============================================================
pptx.layout = "LAYOUT_16x9";
pptx.author = datos.meta.autor;
pptx.company = datos.meta.empresa;
pptx.title = datos.meta.titulo;
pptx.subject = "Oportunidades de Ingeniería y Construcción en Perú";

const m = datos.master;
const W = 10.0;
const H = 5.625;
const MARGIN = 0.5;
const USABLE_W = W - (MARGIN * 2);
const BAR_H = m.barra_superior_altura || 0.4;

// Área útil de contenido (entre la barra superior y el pie)
const CONTENT_TOP = 0.65;
const CONTENT_BOTTOM = 5.18;

// ============================================================
// 2. MASTER SLIDE
// ============================================================
pptx.defineSlideMaster({
  title: "MASTER_TRACADO",
  background: { color: m.fondo },
  objects: [
    {
      rect: { x: 0, y: 0, w: "100%", h: BAR_H, fill: m.barra_superior_color }
    },
    {
      text: {
        text: m.texto_branding,
        options: {
          x: MARGIN, y: 0.08, w: USABLE_W, h: 0.25,
          fontSize: 9,
          color: "FFFFFF",
          fontFace: "Arial",
          align: "left"
        }
      }
    },
    {
      line: {
        x: MARGIN,
        y: H - 0.42,
        w: USABLE_W,
        h: 0,
        line: { color: "CCCCCC", width: 0.5 }
      }
    },
    {
      text: {
        text: m.texto_pie,
        options: {
          x: MARGIN,
          y: H - 0.35,
          w: USABLE_W,
          h: 0.2,
          fontSize: 8,
          color: m.color_texto_pie,
          fontFace: "Arial"
        }
      }
    }
  ]
});

// ============================================================
// 3. FUNCIONES AUXILIARES
// ============================================================

function addTitle(slide, text, y, opts = {}) {
  slide.addText(text, {
    x: MARGIN,
    y: y,
    w: USABLE_W,
    h: opts.h || 0.42,
    fontSize: opts.fontSize || 22,
    bold: true,
    color: m.color_texto_principal,
    fontFace: "Arial",
    ...opts
  });
}

/**
 * Subtítulo con altura 0.35" para permitir 2 líneas de texto.
 * Se posiciona con un pequeño gap debajo del título.
 */
function addSubtitle(slide, text, y) {
  if (!text) return;
  slide.addText(text, {
    x: MARGIN,
    y: y,
    w: USABLE_W,
    h: 0.35,
    fontSize: 11,
    color: m.color_texto_secundario,
    fontFace: "Arial",
    italic: true
  });
}

function addHighlight(slide, text, y, h = 0.5) {
  if (!text) return;
  slide.addText(text, {
    x: MARGIN,
    y: y,
    w: USABLE_W,
    h: h,
    fontSize: 10,
    color: m.color_texto_principal,
    fontFace: "Arial",
    fill: { color: "F0F4FA" },
    align: "center",
    valign: "middle"
  });
}

/**
 * Tags con ancho mínimo garantizado para evitar que textos cortos
 * como "Túneles" o "EPCM" se desborden a 2 líneas.
 */
function addTags(slide, tags, y) {
  if (!tags || tags.length === 0) return;
  let xPos = MARGIN;
  const tagH = 0.24;
  const gap = 0.07;
  tags.forEach((tag) => {
    // ancho proporcional a la longitud, con mínimo de 0.85 pulgadas
    const approxW = Math.max(0.85, (tag.length * 0.075) + 0.16);
    slide.addText(tag, {
      x: xPos,
      y: y,
      w: approxW,
      h: tagH,
      fontSize: 9,
      color: m.color_acento,
      fill: { color: "E8EEF7" },
      align: "center",
      valign: "middle",
      fontFace: "Arial"
    });
    xPos += approxW + gap;
  });
}

// ============================================================
// 4. MOTOR DE RENDERIZADO POR TIPO
// ============================================================

const renderers = {

  portada: (slide, data) => {
    slide.addText(data.titulo, {
      x: MARGIN,
      y: 1.35,
      w: USABLE_W,
      h: 0.85,
      fontSize: 38,
      bold: true,
      color: m.color_texto_principal,
      align: "center",
      fontFace: "Arial"
    });

    slide.addText(data.subtitulo, {
      x: MARGIN,
      y: 2.25,
      w: USABLE_W,
      h: 0.4,
      fontSize: 14,
      color: m.color_texto_secundario,
      align: "center",
      fontFace: "Arial"
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: W / 2 - 0.6,
      y: 2.75,
      w: 1.2,
      h: 0.05,
      fill: m.color_acento
    });

    if (data.kpis && data.kpis.length > 0) {
      const kpiW = (USABLE_W / data.kpis.length) - 0.15;
      const startY = 3.15;
      data.kpis.forEach((kpi, i) => {
        const xPos = MARGIN + (i * (kpiW + 0.15));
        slide.addText(`${kpi.valor}\n${kpi.label}`, {
          x: xPos,
          y: startY,
          w: kpiW,
          h: 1.0,
          fontSize: 12,
          bold: true,
          color: m.color_texto_principal,
          fill: { color: "F5F5F5" },
          align: "center",
          valign: "middle",
          fontFace: "Arial"
        });
      });
    }
  },

  dos_columnas: (slide, data) => {
    addTitle(slide, data.titulo, CONTENT_TOP);
    addSubtitle(slide, data.subtitulo, CONTENT_TOP + 0.48);

    const colW = (USABLE_W - 0.35) / 2;
    const colY = CONTENT_TOP + 0.90;
    const colH = CONTENT_BOTTOM - colY - 0.08;

    if (data.columna_izq && data.columna_izq.length > 0) {
      slide.addText(data.columna_izq.join("\n"), {
        x: MARGIN,
        y: colY,
        w: colW,
        h: colH,
        fontSize: 10.5,
        color: m.color_texto_principal,
        fontFace: "Arial",
        bullet: true,
        lineSpacing: 16,
        paraSpaceAfter: 5,
        valign: "top"
      });
    }

    if (data.columna_der && data.columna_der.length > 0) {
      slide.addText(data.columna_der.join("\n"), {
        x: MARGIN + colW + 0.35,
        y: colY,
        w: colW,
        h: colH,
        fontSize: 10.5,
        color: m.color_texto_principal,
        fontFace: "Arial",
        bullet: true,
        lineSpacing: 16,
        paraSpaceAfter: 5,
        valign: "top"
      });
    }
  },

  kpi_grid: (slide, data) => {
    addTitle(slide, data.titulo, CONTENT_TOP);
    addSubtitle(slide, data.subtitulo, CONTENT_TOP + 0.48);

    if (data.kpis && data.kpis.length > 0) {
      const kpiW = (USABLE_W / data.kpis.length) - 0.15;
      const startY = CONTENT_TOP + 0.95;
      data.kpis.forEach((kpi, i) => {
        const xPos = MARGIN + (i * (kpiW + 0.15));
        const boxColor = kpi.color || m.color_acento;
        slide.addText(`${kpi.valor}\n${kpi.label}`, {
          x: xPos,
          y: startY,
          w: kpiW,
          h: 1.05,
          fontSize: 12,
          bold: true,
          color: "FFFFFF",
          fill: { color: boxColor },
          align: "center",
          valign: "middle",
          fontFace: "Arial"
        });
      });
    }

    if (data.nota) {
      addHighlight(slide, data.nota, CONTENT_BOTTOM - 0.55, 0.5);
    }
  },

  // ---------------------------------------------------------
  // TABLA — gap aumentado entre subtítulo y tabla para evitar
  // que subtítulos largos compriman la tabla hacia abajo.
  // ---------------------------------------------------------
  tabla: (slide, data) => {
    addTitle(slide, data.titulo, CONTENT_TOP);
    // Gap de 0.08" entre título (termina en 1.07) y subtítulo (empieza en 1.13)
    addSubtitle(slide, data.subtitulo, CONTENT_TOP + 0.48);

    if (data.tabla && data.tabla.headers && data.tabla.rows) {
      const numCols = data.tabla.headers.length;
      const colW = USABLE_W / numCols;
      const allRows = [data.tabla.headers, ...data.tabla.rows];
      // Altura proporcional al número de filas, con tope de 2.35"
      const tableH = Math.min(2.35, 0.29 * allRows.length);
      // La tabla empieza 0.10" debajo de donde termina el subtítulo (h=0.35)
      const tableY = CONTENT_TOP + 0.93;

      slide.addTable(allRows, {
        x: MARGIN,
        y: tableY,
        w: USABLE_W,
        h: tableH,
        fontSize: 10,
        fontFace: "Arial",
        color: "333333",
        border: { pt: 0.5, color: "CCCCCC" },
        colW: Array(numCols).fill(colW)
      });

      if (data.tags) {
        addTags(slide, data.tags, tableY + tableH + 0.12);
      }
    }

    if (data.nota) {
      addHighlight(slide, data.nota, CONTENT_BOTTOM - 0.55, 0.5);
    }
  },

  ncc_destacado: (slide, data) => {
    addTitle(slide, data.titulo, CONTENT_TOP);
    addSubtitle(slide, data.subtitulo, CONTENT_TOP + 0.48);

    if (data.kpis && data.kpis.length > 0) {
      const kpiW = (USABLE_W / data.kpis.length) - 0.12;
      const startY = CONTENT_TOP + 0.95;
      data.kpis.forEach((kpi, i) => {
        const xPos = MARGIN + (i * (kpiW + 0.12));
        slide.addText(`${kpi.valor}\n${kpi.label}`, {
          x: xPos,
          y: startY,
          w: kpiW,
          h: 0.55,
          fontSize: 10.5,
          bold: true,
          color: m.color_texto_principal,
          fill: { color: "F5F5F5" },
          align: "center",
          valign: "middle",
          fontFace: "Arial"
        });
      });
    }

    let currentY = CONTENT_TOP + 1.60;
    if (data.tramos && data.tramos.length > 0) {
      data.tramos.forEach((tramo) => {
        slide.addText(`${tramo.nombre}  |  ${tramo.datos}`, {
          x: MARGIN,
          y: currentY,
          w: USABLE_W,
          h: 0.20,
          fontSize: 9.5,
          bold: true,
          color: m.color_acento,
          fontFace: "Arial"
        });
        currentY += 0.23;

        slide.addText(tramo.obras, {
          x: MARGIN + 0.12,
          y: currentY,
          w: USABLE_W - 0.24,
          h: 0.30,
          fontSize: 8.5,
          color: m.color_texto_secundario,
          fontFace: "Arial",
          bullet: true,
          lineSpacing: 14
        });
        currentY += 0.34;
      });
    }

    if (data.magnitud && data.magnitud.length > 0) {
      const magText = data.magnitud.join("  ·  ");
      slide.addText(magText, {
        x: MARGIN,
        y: currentY + 0.04,
        w: USABLE_W,
        h: 0.32,
        fontSize: 8.5,
        color: m.color_texto_principal,
        fontFace: "Arial",
        fill: { color: "FAFAFA" },
        align: "center",
        valign: "middle"
      });
      currentY += 0.40;
    }

    if (data.highlight) {
      addHighlight(slide, data.highlight, CONTENT_BOTTOM - 0.52, 0.48);
    }
  },

  fases: (slide, data) => {
    addTitle(slide, data.titulo, CONTENT_TOP);
    addSubtitle(slide, data.subtitulo, CONTENT_TOP + 0.48);

    if (data.fases && data.fases.length > 0) {
      const availableH = CONTENT_BOTTOM - (CONTENT_TOP + 0.95) - 0.10;
      const boxH = (availableH / data.fases.length) - 0.06;
      let currentY = CONTENT_TOP + 0.95;

      data.fases.forEach((fase) => {
        slide.addShape(pptx.ShapeType.rect, {
          x: MARGIN,
          y: currentY,
          w: USABLE_W,
          h: boxH,
          fill: "FAFAFA",
          line: { color: "E0E0E0", width: 0.5 }
        });

        slide.addText(`${fase.numero}  ·  ${fase.periodo}`, {
          x: MARGIN + 0.1,
          y: currentY + 0.05,
          w: 2.4,
          h: 0.20,
          fontSize: 9,
          bold: true,
          color: m.color_acento,
          fontFace: "Arial"
        });

        slide.addText(fase.titulo, {
          x: MARGIN + 0.1,
          y: currentY + 0.26,
          w: 2.4,
          h: 0.20,
          fontSize: 10.5,
          bold: true,
          color: m.color_texto_principal,
          fontFace: "Arial"
        });

        slide.addText(fase.desc, {
          x: MARGIN + 2.7,
          y: currentY + 0.06,
          w: USABLE_W - 2.9,
          h: boxH - 0.12,
          fontSize: 9.5,
          color: m.color_texto_secundario,
          fontFace: "Arial",
          valign: "middle"
        });

        currentY += boxH + 0.06;
      });
    }
  },

  conclusiones: (slide, data) => {
    addTitle(slide, data.titulo, CONTENT_TOP);
    addSubtitle(slide, data.subtitulo, CONTENT_TOP + 0.48);

    const colW = (USABLE_W - 0.35) / 2;
    const colY = CONTENT_TOP + 0.90;
    // Reservamos 0.58" al final para el highlight
    const colH = (CONTENT_BOTTOM - 0.58) - colY;

    if (data.columna_izq && data.columna_izq.length > 0) {
      slide.addText(data.columna_izq.join("\n"), {
        x: MARGIN,
        y: colY,
        w: colW,
        h: colH,
        fontSize: 10.5,
        color: m.color_texto_principal,
        fontFace: "Arial",
        bullet: true,
        lineSpacing: 16,
        paraSpaceAfter: 5,
        valign: "top"
      });
    }

    if (data.columna_der && data.columna_der.length > 0) {
      slide.addText(data.columna_der.join("\n"), {
        x: MARGIN + colW + 0.35,
        y: colY,
        w: colW,
        h: colH,
        fontSize: 10.5,
        color: m.color_texto_principal,
        fontFace: "Arial",
        bullet: true,
        lineSpacing: 16,
        paraSpaceAfter: 5,
        valign: "top"
      });
    }

    if (data.highlight) {
      addHighlight(slide, data.highlight, CONTENT_BOTTOM - 0.53, 0.48);
    }
  }
};

// ============================================================
// 5. ORQUESTADOR PRINCIPAL
// ============================================================

console.log(`📊 Generando presentación: ${datos.meta.titulo}`);
console.log(`📑 Total de slides en JSON: ${datos.slides.length}\n`);

datos.slides.forEach((slideData, index) => {
  const tipo = slideData.tipo;
  const renderer = renderers[tipo];

  if (!renderer) {
    console.warn(`⚠️  Tipo de slide desconocido: "${tipo}" — se omite slide ${index + 1}`);
    return;
  }

  const slide = pptx.addSlide({ masterName: "MASTER_TRACADO" });
  renderer(slide, slideData);
  console.log(`✅ Slide ${index + 1}: ${tipo.padEnd(18)} | ${slideData.titulo}`);
});

// ============================================================
// 6. EXPORTACIÓN
// ============================================================

const outputName = "Tracado_Peru_2026-2035.pptx";

pptx.writeFile({ fileName: outputName })
  .then(() => {
    console.log(`\n🎉 Archivo generado exitosamente: ${outputName}`);
    console.log("📂 Abre el archivo en PowerPoint o LibreOffice Impress para revisar.");
  })
  .catch((err) => {
    console.error("❌ Error al generar el archivo:", err);
    process.exit(1);
  });
