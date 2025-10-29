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

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');

// Parámetros de diseño de la tarjeta

const CARD = {
  width: 280,
  height: 95,
  gapX: 20,
  gapY: 15,
  margin: 15
};

async function main() {
  const [,, inputFile, marcoFile, outputPdf] = process.argv;
  if (!inputFile || !marcoFile || !outputPdf) {
    console.error('Uso: node generador.js <input.xlsx> <marco.png> <output.pdf>');
    process.exit(1);
  }

  // Leer Excel y parsear
  const workbook = xlsx.readFile(inputFile);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });
  if (!rows.length) {
    console.error('El archivo Excel está vacío.');
    process.exit(1);
  }

  // Detectar columnas dinámicamente
  const headers = Object.keys(rows[0]);
  let productoKey = headers.find(h=>/producto/i.test(h));
  let precioKey  = headers.find(h=>/precio/i.test(h));
  if (!productoKey || !precioKey) {
    if (headers.length===2) [productoKey,precioKey]=headers;
    else { console.error('Encabezados inválidos.'); process.exit(1); }
  }

  // Normalizar datos
  const data = rows.map(r=>({
    product: String(r[productoKey]).toUpperCase(),
    price: String(r[precioKey]).toUpperCase()
  }));

  await generatePdf(data, outputPdf, marcoFile);
  console.log(`PDF generado: ${outputPdf}`);
}

/**
 * generatePdf: genera el PDF con tarjetas y líneas de recorte
 */
function generatePdf(data, outputPdf, marcoFile) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size:'LETTER', margin:CARD.margin });
    // Registrar Arial
    doc.registerFont('Arial', path.join(__dirname,'Arial.ttf'));
    doc.registerFont('Arial-Bold', path.join(__dirname,'Arial-Bold.ttf'));

    const stream = fs.createWriteStream(outputPdf);
    doc.pipe(stream);

    const pageW = doc.page.width, pageH = doc.page.height;
    const columns = Math.floor((pageW - 2*CARD.margin + CARD.gapX)/(CARD.width+CARD.gapX));
    const rowsCount = Math.floor((pageH-2*CARD.margin+CARD.gapY)/(CARD.height+CARD.gapY));
    const perPage = columns*rowsCount;

    // Páginas
    for(let p=0; p*perPage<data.length; p++){
      if(p>0) doc.addPage();
      const pageItems = data.slice(p*perPage,p*perPage+perPage);
      // Dibujar tarjetas
      pageItems.forEach((item,i)=>{
        const c = i%columns, r = Math.floor(i/columns);
        const x = CARD.margin+c*(CARD.width+CARD.gapX);
        const y = CARD.margin+r*(CARD.height+CARD.gapY);

        // --- PASO 1: DIBUJAR EL FONDO/MARCO ---
        try {
          doc.image(marcoFile, x, y, {
            width: CARD.width,
            height: CARD.height
          });
        } catch (e) {
          console.error(`No se pudo cargar la imagen del marco: ${marcoFile}`);
          doc.save().lineWidth(1).strokeColor('red')
             .rect(x,y,CARD.width,CARD.height).stroke().restore();
        }

        // --- PASO 2: DIBUJAR EL TEXTO (ENCIMA DEL FONDO) ---

        // Definir área de texto
        const padL=10, padR=10, spacing=4;
        const tx = x + padL;
        const tw = CARD.width - (padL + padR);

        // 1. Calcular tamaño de PRECIO (Arial-Bold, 14pt max)
        let precioSize=14, precioHeight;
        for(let sz=14; sz>=6; sz--){
          doc.font('Arial-Bold').fontSize(sz);
          precioHeight = doc.heightOfString(item.price,{width:tw,align:'center'});
          if(precioHeight <= sz*1.2*2){ precioSize=sz; break; }
        }
        doc.font('Arial-Bold').fontSize(precioSize);
        precioHeight = doc.heightOfString(item.price,{width:tw,align:'center'});

        // 2. Calcular tamaño de PRODUCTO (Arial regular, 10pt max)
        let productoSize=10, productoHeight;
        for(let sz=10; sz>=6; sz--){
          doc.font('Arial').fontSize(sz);
          productoHeight = doc.heightOfString(item.product,{width:tw,align:'center'});
          if(productoHeight <= sz*1.2*2){ productoSize=sz; break; }
        }
        doc.font('Arial').fontSize(productoSize);
        productoHeight = doc.heightOfString(item.product,{width:tw,align:'center'});

        // 3. Centrar y Dibujar (Producto primero, luego Precio)
        const totalHeight = productoHeight + spacing + precioHeight;
        const ty = y + (CARD.height - totalHeight) / 2;
        
        // ¡¡AQUÍ ESTÁ LA MAGIA!! Fijamos el color ANTES de dibujar
        // Cambia 'white' por 'black' si tu fondo es claro
        doc.fillColor('black'); 

        // Dibujar PRODUCTO (Arriba, regular, 10pt)
        doc.font('Arial').fontSize(productoSize)
           .text(item.product, tx, ty, {width:tw, align:'center'});
        
        // Dibujar PRECIO (Abajo, bold, 14pt)
        doc.font('Arial-Bold').fontSize(precioSize)
           .text(item.price, tx, ty + productoHeight + spacing, {width:tw, align:'center'});
     });
      // líneas de recorte
      doc.save().lineWidth(0.5).strokeColor('#999').dash(5,{space:5});
      // verticales
      for(let c=1;c<columns;c++){
        const xL=CARD.margin+c*(CARD.width+CARD.gapX)-CARD.gapX/2;
        doc.moveTo(xL,CARD.margin-5).lineTo(xL,pageH-CARD.margin+5).stroke();
      }
      // horizontales
      for(let r=1;r<rowsCount;r++){
        const yL=CARD.margin+r*(CARD.height+CARD.gapY)-CARD.gapY/2;
        doc.moveTo(CARD.margin-5,yL).lineTo(pageW-CARD.margin+5,yL).stroke();
      }
      doc.undash().restore();
    }

    doc.end();
    stream.on('finish',resolve);
    stream.on('error',reject);
  });
}

main().catch(e=>{console.error('Error:',e);process.exit(1);});
