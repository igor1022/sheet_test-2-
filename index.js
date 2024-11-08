import { google } from 'googleapis';
import readlineSync from 'readline-sync';
import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';  // Импортируем библиотеку для работы с Excel

// OAuth 2.0 параметры (из вашего Google Cloud Console)
const CLIENT_ID = '385891869614-rb1225grm5vro613vsnvg56a0vtm500s.apps.googleusercontent.com';
const CLIENT_SECRET = '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';
const REDIRECT_URIS = ['http://localhost:3000/oauth2callback'];

// Идентификатор вашего Google Sheets документа
const spreadsheetId = '1H8t979DnErL0xxtA0TLQadW0jXy-5Eb8uVqxwKY88gw'; // Замените на ваш ID документа
const range = 'Ответы на форму (1)!A:N'; // Задаем диапазон с A по N столбец

// Путь для хранения токена
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const OUTPUT_EXCEL_PATH = path.join(process.cwd(), 'output.xlsx'); // Путь для сохранения Excel файла

// Функция для авторизации
async function authorize() {
  const { client_secret, client_id, redirect_uris } = { client_secret: CLIENT_SECRET, client_id: CLIENT_ID, redirect_uris: REDIRECT_URIS };
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Проверяем, есть ли уже сохраненный токен
  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  // Если нет токена, инициируем процесс авторизации
  return getNewToken(oAuth2Client);
}

// Функция для получения нового токена
function getNewToken(oAuth2Client) {
  return new Promise((resolve, reject) => {
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
    });

    console.log('Authorize this app by visiting this url:', authUrl);
    const code = readlineSync.question('Enter the code from the page here: ');

    oAuth2Client.getToken(code, (err, tokens) => {
      if (err) {
        return reject('Error retrieving access token: ' + err);
      }
      oAuth2Client.setCredentials(tokens);
      fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens)); // Сохраняем токен
      resolve(oAuth2Client);
    });
  });
}

// Функция для чтения данных из Google Sheets
async function getSheetData(auth) {
  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const rows = response.data.values;
    if (rows.length) {
      console.log('Data retrieved successfully.');

      // Создаем новый рабочий лист, используя только данные с A до N столбца
      const ws = XLSX.utils.aoa_to_sheet(rows);

      // Создаем книгу Excel и добавляем рабочий лист
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Ответы на форму (1)');

      // Сохраняем книгу в файл Excel
      XLSX.writeFile(wb, OUTPUT_EXCEL_PATH);
      console.log(`Data saved to ${OUTPUT_EXCEL_PATH}`);
    } else {
      console.log('No data found.');
    }
  } catch (error) {
    console.error('The API returned an error: ' + error);
  }
}

// Основная функция
async function main() {
  try {
    const auth = await authorize();
    await getSheetData(auth);
  } catch (error) {
    console.error('Error retrieving sheet data:', error);
  }
}

main();
