// generador.js (Modificado para ser un módulo)

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const PDFDocument = require("pdfkit");

// --- Tus funciones de ayuda (sin cambios) ---
const CARD = {
  width: 280,
  height: 95,
  gapX: 20,
  gapY: 15,
  margin: 15,
  topPadding: 15,
};

function normalizeUnitsInName(name) {
  if (!name) return name;
  const units = ["GR", "KG", "LT", "ML"];
  const re = new RegExp(
    "(\\d{1,3}(?:[.,]\\d{3})*(?:[.,]\\d+)?)\\s*(" + units.join("|") + ")\\b",
    "gi"
  );
  return String(name).replace(re, (match, numStr, unit) => {
    let tmp = numStr;
    if (/\.\d{3},/.test(tmp)) {
      tmp = tmp.replace(/\./g, "").replace(/,/g, ".");
    } else {
      tmp = tmp.replace(/\.(?=\d{3}(?:[.,]|$))/g, "");
      tmp = tmp.replace(",", ".");
    }
    const num = parseFloat(tmp);
    if (isNaN(num)) return match;
    let formatted;
    if (Math.round(num * 100) % 100 === 0) {
      formatted = String(Math.round(num));
    } else {
      formatted = (Math.round(num * 100) / 100).toString().replace(/\.?0+$/, "");
    }
    return formatted + unit.toUpperCase();
  });
}

function formatPrice(value) {
  if (value == null) return "";
  let str = String(value).trim();
  str = str.replace(/[^\d,.\-]/g, "");
  str = str.replace(/\.(?=\d{3}(?:[.,]|$))/g, "");
  str = str.replace(",", ".");
  const num = parseFloat(str);
  if (isNaN(num)) return "";
  const fixedStr = num.toFixed(2);
  const parts = fixedStr.split(".");
  let integerPart = parts[0];
  const decimalPart = parts[1];
  integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return integerPart + "," + decimalPart;
}
// --- Fin de tus funciones de ayuda ---


/**
 * ¡NUEVA FUNCIÓN EXPORTABLE!
 * Esta es la función que api/bot.js llamará.
 * Reemplaza la antigua lógica de 'main'.
 */
async function generatePricePdf(excelFilePath, marcoFilePath, outputPdfPath) {
  // Leer Excel y parsear
  const workbook = xlsx.readFile(excelFilePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows.length) {
    throw new Error("El archivo Excel está vacío.");
  }

  // Detectar columnas dinámicamente
  const headers = Object.keys(rows[0]);
  let productoKey = headers.find((h) => /producto/i.test(h));
  let precioKey = headers.find((h) => /precio/i.test(h));
  if (!productoKey || !precioKey) {
    if (headers.length === 2) [productoKey, precioKey] = headers;
    else {
      throw new Error("Encabezados inválidos.");
    }
  }

  // --- Construcción del array final ---
  const data = rows.map((r) => {
    const rawName = String(r[productoKey] || "");
    const cleanedName = normalizeUnitsInName(rawName).toUpperCase();
    const cleanedPrice = formatPrice(r[precioKey]);
    return {
      product: cleanedName,
      price: `BS ${cleanedPrice}`,
    };
  });

  // Llamar a tu lógica de generación de PDF
  await generatePdf(data, outputPdfPath, marcoFilePath);
  console.log(`PDF generado: ${outputPdfPath}`);
}


/**
 * Tu lógica de PDFKit (casi sin cambios)
 */
function generatePdf(data, outputPdf, marcoFile) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "LETTER", margin: CARD.margin });
    
    // --- ¡CAMBIO IMPORTANTE! ---
    // Registramos la fuente usando __dirname para que Vercel la encuentre
    doc.registerFont("Arial", path.join(__dirname, "Altone Trial-Bold.ttf"));
    doc.registerFont("Arial-Bold", path.join(__dirname, "Altone Trial-Bold.ttf"));
    // ---------------------------

    const stream = fs.createWriteStream(outputPdf);
    doc.pipe(stream);

    const pageW = doc.page.width,
      pageH = doc.page.height;
    const columns = Math.floor(
      (pageW - 2 * CARD.margin + CARD.gapX) / (CARD.width + CARD.gapX)
    );
    const rowsCount = Math.floor(
      (pageH - 2 * CARD.margin + CARD.gapY) / (CARD.height + CARD.gapY)
    );
    const perPage = columns * rowsCount;

    // Páginas
    for (let p = 0; p * perPage < data.length; p++) {
      if (p > 0) doc.addPage();
      const pageItems = data.slice(p * perPage, p * perPage + perPage);
      // Dibujar tarjetas
      pageItems.forEach((item, i) => {
        const c = i % columns,
          r = Math.floor(i / columns);
        const x = CARD.margin + c * (CARD.width + CARD.gapX);
        const y = CARD.margin + r * (CARD.height + CARD.gapY);

        // PASO 1: DIBUJAR EL FONDO/MARCO
        try {
          doc.image(marcoFile, x, y, {
            width: CARD.width,
            height: CARD.height,
          });
        } catch (e) {
          console.error(`No se pudo cargar la imagen del marco: ${marcoFile}`);
          doc
            .save()
            .lineWidth(1)
            .strokeColor("red")
            .rect(x, y, CARD.width, CARD.height)
            .stroke()
            .restore();
        } 

        // PASO 2: DIBUJAR EL TEXTO (Tu lógica sin cambios)
        const padR = 10, spacing = -3;
        const padL = CARD.width / 2.9;
        const tx = x + padL;
        const tw = CARD.width - padL - padR;

        doc.fillColor("#545454"); 
        let productoSize = 20, productoHeight;
        for (let sz = 20; sz >= 7; sz--) {
          doc.font("Arial").fontSize(sz);
          productoHeight = doc.heightOfString(item.product, {
            width: tw,
            align: "center",
            lineGap: -1,
          });
          if (productoHeight <= sz * 1.2 * 3) {
            productoSize = sz;
            break;
          }
        }
        doc.font("Arial").fontSize(productoSize);
        productoHeight = doc.heightOfString(item.product, {
          width: tw,
          align: "center",
          lineGap: -1,
        });

        let fullPrice = item.price;
        let integerPart = fullPrice;
        let decimalPart = "";
        if (fullPrice.includes(",")) {
          const parts = fullPrice.split(",");
          integerPart = parts[0] + ",";
          decimalPart =
            parts[1].length > 2 ? parts[1].substring(0, 2) : parts[1];
        }

        let precioSize = 28, precioHeight;
        for (let sz = 28; sz >= 7; sz--) {
          doc.font("Arial-Bold").fontSize(sz);
          precioHeight = doc.heightOfString(integerPart, {
            width: tw,
            align: "center",
            lineGap: -1,
          });
          if (precioHeight <= sz * 1.2 * 2) {
            precioSize = sz;
            break;
          }
        }
        doc.font("Arial-Bold").fontSize(precioSize);
        precioHeight = doc.heightOfString(integerPart, {
          width: tw,
          align: "center",
          lineGap: -1,
        });
        
        const decimalSize = Math.max(8, precioSize - 6);
        doc.font("Arial-Bold").fontSize(precioSize);
        const integerWidth = doc.widthOfString(integerPart);
        doc.font("Arial-Bold").fontSize(decimalSize);
        const decimalWidth = doc.widthOfString(decimalPart);
        const totalPrecioWidth = integerWidth + decimalWidth;
        const totalHeight = precioHeight + spacing + productoHeight;
        const ty = y + (CARD.height - totalHeight) / 2;
        const precioStartX = tx + (tw - totalPrecioWidth) / 2;
        
        doc
          .font("Arial-Bold")
          .fontSize(precioSize)
          .fillColor("#545454")
          .text(integerPart, precioStartX, ty, {
            lineBreak: false,
            lineGap: -1,
          });
        
        doc
          .font("Arial-Bold")
          .fontSize(decimalSize)
          .fillColor("#545454")
          .text(decimalPart, precioStartX + integerWidth, ty, {
            lineBreak: false,
          });
        
        const productoY = ty + precioHeight + spacing;
        doc
          .font("Arial")
          .fontSize(productoSize)
          .fillColor("#545454")
          .text(item.product, tx, productoY, {
            width: tw,
            align: "center",
            lineGap: -1,
          });
      });
      
      // líneas de recorte (sin cambios)
      doc.save().lineWidth(0.5).strokeColor("#999").dash(5, { space: 5 });
      for (let c = 1; c < columns; c++) {
        const xL = CARD.margin + c * (CARD.width + CARD.gapX) - CARD.gapX / 2;
        doc
          .moveTo(xL, CARD.margin - 5)
          .lineTo(xL, pageH - CARD.margin + 5)
          .stroke();
      }
      for (let r = 1; r < rowsCount; r++) {
        const yL = CARD.margin + r * (CARD.height + CARD.gapY) - CARD.gapY / 2;
        doc
          .moveTo(CARD.margin - 5, yL)
          .lineTo(pageW - CARD.margin + 5, yL)
          .stroke();
      }
      doc.undash().restore();
    }

    doc.end();
    stream.on("finish", resolve);
    stream.on("error", reject);
  });
}

// ¡Lo más importante! Exportamos la función para que bot.js la use.
module.exports = { generatePricePdf };