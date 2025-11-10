// api/bot.js (El cerebro del bot con Telegraf)

const { Telegraf, Markup } = require('telegraf');
const fs = require('fs');
const path = require('path');
const https = require('https');
// Importamos NUESTRO generador de PDF
const { generatePricePdf } = require('../generador.js'); 

// Constantes
const TEMP_DIR = '/tmp'; // Único directorio escribible en Vercel
const TOKEN = process.env.TELEGRAM_TOKEN; // Token desde Vercel
const bot = new Telegraf(TOKEN);

// --- Funciones de Ayuda ---

// Función para descargar archivos de Telegram
function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    https.get(url, (response) => {
      response.pipe(file);
      file.on('finish', () => file.close(resolve));
    }).on('error', (err) => fs.unlink(dest, () => reject(err)));
  });
}

// Función para el nombre de archivo (la adapté de tu index.js)
function getTimestampFilename() {
  const pad = n => String(n).padStart(2, '0');
  const now = new Date();
  // Usamos la zona horaria de Vercel (UTC) por simplicidad, 
  // o puedes re-implementar la de Venezuela si prefieres.
  const ts = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}`;
  return `indicador-precios-${ts}.pdf`;
}

// --- Lógica del Bot ---

// 1. Manejador de /start o cualquier texto
const startHandler = (ctx) => {
  ctx.reply(
    '👋 Soy el Bot de Generador de Indicadores de precios. Selecciona una opción:',
    Markup.inlineKeyboard([
      // Cada botón en su propio array para apilarlos
      [ Markup.button.callback('Generar indicadores', 'GENERATE') ],
      [ Markup.button.callback('Ayuda', 'HELP') ]
    ])
  );
};

bot.start(startHandler);
bot.on('text', startHandler); // Responde igual a cualquier texto

// 2. Manejador del botón "Ayuda"
bot.action('HELP', (ctx) => {
  ctx.answerCbQuery(); // Quita el "cargando" del botón
  ctx.reply(
    '📖 *Ayuda*: Presiona "Generar indicadores de precios" para comenzar. ' +
    'Luego, envíame el archivo Excel con las columnas "producto" y "precio".',
    { parse_mode: 'Markdown' }
  );
});

// 3. Manejador del botón "Generar"
bot.action('GENERATE', (ctx) => {
  ctx.answerCbQuery();
  ctx.reply('📂 Por favor, envía un archivo Excel (.xlsx) con las columnas "producto" y "precio".');
});

// 4. Manejador de Documentos (Aquí ocurre la magia)
bot.on('document', async (ctx) => {
  const chatId = ctx.chat.id;
  const doc = ctx.message.document;

  // 4.1. Verificar que sea un Excel
  if (doc.mime_type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' && !doc.file_name.endsWith('.xlsx')) {
    return ctx.reply('❌ Ese no es un archivo Excel. Por favor, envía un archivo `.xlsx`.');
  }

  try {
    // 4.2. Informar al usuario
    await ctx.reply('⏳ Procesando tu archivo Excel... Esto puede tardar un momento.');

    // 4.3. Definir todos los paths
    const excelLink = await ctx.telegram.getFileLink(doc.file_id);
    const excelPath = path.join(TEMP_DIR, `input_${chatId}.xlsx`);
    const outputPath = path.join(TEMP_DIR, `output_${chatId}.pdf`);
    
    // ¡IMPORTANTE! Path al archivo estático 'marco.png'
    // __dirname es /api, '..' sube a la raíz del proyecto
    const marcoPath = path.join(__dirname, '..', 'marco.png');

    // 4.4. Descargar el Excel
    await downloadFile(excelLink.href, excelPath);

    // 4.5. Generar el nombre de archivo final
    const finalFilename = getTimestampFilename();

    // 4.6. ¡Ejecutar nuestra lógica de PDF!
    // No más 'exec', llamamos a la función directamente
    await generatePricePdf(excelPath, marcoPath, outputPath);

    // 4.7. Enviar el PDF de vuelta
    await ctx.replyWithDocument(
      { source: outputPath, filename: finalFilename },
      { caption: '¡Aquí tienes tus indicadores de precios! ✨' }
    );

    // 4.8. Limpiar archivos temporales
    fs.unlinkSync(excelPath);
    fs.unlinkSync(outputPath);

  } catch (err) {
    console.error('Error al procesar documento:', err);
    await ctx.reply('❌ Ocurrió un error al procesar tu archivo. Asegúrate de que las columnas se llamen "producto" y "precio".');
  }
});

// --- Función Serverless de Vercel ---
module.exports = async (request, response) => {
  try {
    await bot.handleUpdate(request.body);
  } catch (err) {
    console.error('Error al manejar el update:', err);
  }
  response.status(200).send('OK');
};