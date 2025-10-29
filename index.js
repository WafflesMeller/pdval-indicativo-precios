// index.js

// 1. Cargar variables de entorno
require('dotenv').config();

// 2. Importar librerías necesarias
const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');
const { exec } = require('child_process');

// 3. Token y endpoint de Telegram
const BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_API = `https://api.telegram.org/bot${BOT_TOKEN}`;

// 4. Crear aplicación de Express
const app = express();
app.use(express.json());

// 5. Webhook para recibir actualizaciones
app.post('/webhook', async (req, res) => {
  console.log('Update recibido:', req.body);
  res.sendStatus(200);

  const update = req.body;

  // 5.1. Manejar callback queries (botones)
  if (update.callback_query) {
    const { id: callbackQueryId, data, message } = update.callback_query;
    const chatId = message.chat.id;
    // Quitar spinner
    await axios.post(`${TELEGRAM_API}/answerCallbackQuery`, { callback_query_id: callbackQueryId });

    if (data === 'GENERATE') {
      // Solicitar el archivo Excel
      await axios.post(`${TELEGRAM_API}/sendMessage`, {
        chat_id: chatId,
        text: '📂 Por favor, envía un archivo Excel (.xlsx) con dos columnas: "nombre" y "cargo".'
      });
    } else if (data === 'HELP') {
      // Ayuda básica
      await axios.post(`${TELEGRAM_API}/sendMessage`, {
        chat_id: chatId,
        text: '📖 *Ayuda*: Presiona "Generar precedencias" para comenzar.',
        parse_mode: 'Markdown'
      });
    }
    return;
  }

  // 5.2. Manejar documentos Excel (.xlsx)
  if (update.message && update.message.document && update.message.document.mime_type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    const chatId = update.message.chat.id;
    try {
      // Informar procesamiento
      await axios.post(`${TELEGRAM_API}/sendMessage`, { chat_id: chatId, text: '⏳ Procesando tu archivo Excel...' });

      // Descargar el archivo
      const fileId = update.message.document.file_id;
      const fileInfo = await axios.get(`${TELEGRAM_API}/getFile?file_id=${fileId}`);
      const filePath = fileInfo.data.result.file_path;
      const downloadUrl = `https://api.telegram.org/file/bot${BOT_TOKEN}/${filePath}`;
      const excelRes = await axios.get(downloadUrl, { responseType: 'arraybuffer' });

      // Guardar localmente
      const inputPath = path.join(__dirname, 'input.xlsx');
      fs.writeFileSync(inputPath, excelRes.data);

      // Preparar rutas
      const logoPath = path.join(__dirname, 'logo.png');
      const outputPdf = path.join(__dirname, 'precedencias.pdf');

      // Ejecutar el script generador
      exec(`node generador.js "${inputPath}" "${logoPath}" "${outputPdf}"`, async (error, stdout, stderr) => {
        if (error) {
          console.error('Error generando PDF:', stderr);
          return axios.post(`${TELEGRAM_API}/sendMessage`, { chat_id: chatId, text: '❌ Error al generar las precedencias.' });
        }

        // Enviar PDF generado
        const form = new FormData();
        form.append('chat_id', chatId);
        form.append('document', fs.createReadStream(outputPdf));
        await axios.post(`${TELEGRAM_API}/sendDocument`, form, { headers: form.getHeaders() });

        // Limpiar archivos temporales
        fs.unlinkSync(inputPath);
        fs.unlinkSync(outputPdf);
      });
    } catch (err) {
      console.error('Error al procesar documento:', err);
      await axios.post(`${TELEGRAM_API}/sendMessage`, { chat_id: chatId, text: '❌ Ocurrió un error al procesar tu archivo.' });
    }
    return;
  }

  // 5.3. Manejar mensajes de texto normales
  if (update.message && update.message.text) {
    const chatId = update.message.chat.id;
    await axios.post(`${TELEGRAM_API}/sendMessage`, {
      chat_id: chatId,
      text: '👋 Soy el Bot de Generador de Precedencias. Selecciona una opción:',
      reply_markup: {
        inline_keyboard: [
          [
            { text: 'Generar precedencias', callback_data: 'GENERATE' },
            { text: 'Ayuda', callback_data: 'HELP' }
          ]
        ]
      }
    });
  }
});

// 6. Iniciar servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en puerto ${PORT}`));
