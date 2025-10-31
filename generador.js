// generador.js

/**
 * Generador de precedencias - generador.js
 *
 * Este script:
 *   1) Lee un archivo Excel (.xlsx/.xls) con columnas “nombre” y “cargo”.
 *   2) Genera un PDF con tarjetas formateadas utilizando PDFKit y fuentes Arial.
 *   3) Añade líneas de recorte entre filas y columnas que sobresalgan del borde.
 *
 * Uso:
 *   node generador.js <input.xlsx> <logo.png> <output.pdf>
 */

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const PDFDocument = require("pdfkit");

// Parámetros de diseño de la tarjeta

const CARD = {
  width: 280,
  height: 95,
  gapX: 20,
  gapY: 15,
  margin: 15,
  topPadding: 15,
};

async function main() {
  const [, , inputFile, marcoFile, outputPdf] = process.argv;
  if (!inputFile || !marcoFile || !outputPdf) {
    console.error(
      "Uso: node generador.js <input.xlsx> <marco.png> <output.pdf>"
    );
    process.exit(1);
  }

  // Leer Excel y parsear
  const workbook = xlsx.readFile(inputFile);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows.length) {
    console.error("El archivo Excel está vacío.");
    process.exit(1);
  }

  // Detectar columnas dinámicamente
  const headers = Object.keys(rows[0]);
  let productoKey = headers.find((h) => /producto/i.test(h));
  let precioKey = headers.find((h) => /precio/i.test(h));
  if (!productoKey || !precioKey) {
    if (headers.length === 2) [productoKey, precioKey] = headers;
    else {
      console.error("Encabezados inválidos.");
      process.exit(1);
    }
  }

  // Normalizar datos
  const data = rows.map((r) => {
    const name = String(r[productoKey] || "").toUpperCase();

    // 🔹 Función para limpiar y formatear el precio
    function formatPrice(value) {
      if (value == null) return "";

      let str = String(value).trim();

      // 1️⃣ Eliminar cualquier símbolo de moneda o espacio extra
      str = str.replace(/[^\d,.\-]/g, "");

      // 2️⃣ Detectar el formato y normalizar:
      // - Eliminar puntos que parecen separadores de miles
      // - Convertir coma en punto decimal
      //   Ej: "1.000,50" -> "1000.50", "1,000.50" -> "1000.50"
      str = str.replace(/\.(?=\d{3}(?:[.,]|$))/g, ""); // borra puntos de miles
      str = str.replace(",", "."); // cambia coma a punto decimal

      const num = parseFloat(str);
      if (isNaN(num)) return "";

      // 3️⃣ Formatear con separador de miles "." y decimales ","
      const formatted = num.toLocaleString("es-ES", {
        minimumFractionDigits: 2, // siempre 2 decimales
        maximumFractionDigits: 2,
      });

      return formatted;
    }

    // 🔹 Aplicar limpieza del precio
    const cleanedPrice = formatPrice(r[precioKey]);

    return {
      product: name,
      price: `BS ${cleanedPrice}`,
    };
  });

  await generatePdf(data, outputPdf, marcoFile);
  console.log(`PDF generado: ${outputPdf}`);
}

/**
 * generatePdf: genera el PDF con tarjetas y líneas de recorte
 */
function generatePdf(data, outputPdf, marcoFile) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "LETTER", margin: CARD.margin });
    // Registrar Arial
    doc.registerFont("Arial", path.join(__dirname, "Altone Trial-Bold.ttf"));
    doc.registerFont(
      "Arial-Bold",
      path.join(__dirname, "Altone Trial-Bold.ttf")
    );

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

        // --- PASO 1: DIBUJAR EL FONDO/MARCO ---
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
        } // Definir área de texto con 1/3 de margen izquierdo

        // --- PASO 2: DIBUJAR EL TEXTO (ENCIMA DEL FONDO) ---

        const padR = 10,
          spacing = -3;
        const padL = CARD.width / 2.9;
        const tx = x + padL;
        const tw = CARD.width - padL - padR;

        // Fijar color (ajusta 'black' o 'white' según tu fondo)
        doc.fillColor("#545454"); // 1. Calcular tamaño de PRODUCTO (Arial regular, 20pt max)

        let productoSize = 20,
          productoHeight;
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

        // 2. Dividir el Precio en parte entera y decimal
        let fullPrice = item.price;
        let integerPart = fullPrice;
        let decimalPart = "";
        if (fullPrice.includes(".")) {
          const parts = fullPrice.split(".");
          integerPart = parts[0] + "."; // "BS 123."
          decimalPart =
            parts[1].length > 2 ? parts[1].substring(0, 2) : parts[1]; // "45" (máx 2 decimales)
        } // 3. Calcular tamaño de PRECIO (Solo parte entera, Arial-Bold, 28pt max)

        let precioSize = 28,
          precioHeight;
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

        // 4. Calcular tamaño y ancho de los decimales
        const decimalSize = Math.max(8, precioSize - 4); // 4pt más pequeño que el precio
        doc.font("Arial-Bold").fontSize(precioSize);
        const integerWidth = doc.widthOfString(integerPart);
        doc.font("Arial-Bold").fontSize(decimalSize);
        const decimalWidth = doc.widthOfString(decimalPart);
        const totalPrecioWidth = integerWidth + decimalWidth; // 5. Centrar Verticalmente

        const totalHeight = productoHeight + spacing + precioHeight;
        const ty = y + (CARD.height - totalHeight) / 2;

        // 6. DIBUJAR PRODUCTO (Centrado en el área de texto)
        doc
          .font("Arial")
          .fontSize(productoSize)
          .fillColor("#545454") // <-- Asegúrate que el color sea visible
          .text(item.product, tx, ty, {
            width: tw,
            align: "center",
            lineGap: -1,
          });

        // 7. DIBUJAR PRECIO (Centrado manual de las dos partes)
        const precioY = ty + productoHeight + spacing;
        // Calcular el 'X' inicial para centrar ambas partes juntas
        const precioStartX = tx + (tw - totalPrecioWidth) / 2;

        // Parte Entera
        doc
          .font("Arial-Bold")
          .fontSize(precioSize)
          .fillColor("#545454") // <-- Asegúrate que el color sea visible
          .text(integerPart, precioStartX, precioY, {
            lineBreak: false,
            lineGap: -1,
          });

        // Parte Decimal (mismo 'Y' para alinear por arriba)
        doc
          .font("Arial-Bold")
          .fontSize(decimalSize)
          .fillColor("#545454") // <-- Asegúrate que el color sea visible
          .text(decimalPart, precioStartX + integerWidth, precioY, {
            lineBreak: false,
          });
      });
      // líneas de recorte
      doc.save().lineWidth(0.5).strokeColor("#999").dash(5, { space: 5 });
      // verticales
      for (let c = 1; c < columns; c++) {
        const xL = CARD.margin + c * (CARD.width + CARD.gapX) - CARD.gapX / 2;
        doc
          .moveTo(xL, CARD.margin - 5)
          .lineTo(xL, pageH - CARD.margin + 5)
          .stroke();
      }
      // horizontales
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

main().catch((e) => {
  console.error("Error:", e);
  process.exit(1);
});
