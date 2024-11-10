import { google } from 'googleapis';
import readlineSync from 'readline-sync';
import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';  // Импортируем библиотеку для работы с Excel
import sql from 'mssql';

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

// Настройки подключения к базе данных
const dbConfig = {
    user: 'dbadmin',
    password: 'Account8',
    server: 'sheetstest.database.windows.net',
    database: 'master',
    options: {
        encrypt: true,
        trustServerCertificate: false
    }
};

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

// Функция для чтения данных из базы данных
async function getDbData() {
  try {
    await sql.connect(dbConfig);
    console.log('Подключение к базе данных успешно!');

    // Выполнение SQL запроса
    const result = await sql.query('SELECT TOP 10 * FROM your_table_name');
    console.log('Данные из базы данных:', result.recordset);

    return result.recordset;
  } catch (err) {
    console.error('Ошибка при подключении к базе данных:', err);
  }
}

// Функция для записи данных в Google Sheets
async function updateSheetData(auth, data) {
  const sheets = google.sheets({ version: 'v4', auth });

  try {
    // Подготовка данных для записи в Google Sheets
    const values = data.map(row => [
      row.Column1,  // Замените на реальные имена столбцов из вашей таблицы
      row.Column2
    ]);

    // Запись данных в Google Sheets
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'Sheet1!A1', // Укажите диапазон
      valueInputOption: 'RAW',
      requestBody: {
        values
      }
    });

    console.log('Данные успешно записаны в Google Sheets');
  } catch (error) {
    console.error('Ошибка при записи данных в Google Sheets:', error);
  }
}

// Основная функция
async function main() {
  try {
    // Получаем данные из базы данных
    const dbData = await getDbData();

    if (!dbData) {
      console.log('Нет данных для записи');
      return;
    }

    // Авторизуемся в Google Sheets
    const auth = await authorize();

    // Обновляем данные в Google Sheets
    await updateSheetData(auth, dbData);
  } catch (error) {
    console.error('Ошибка:', error);
  }
}

// Вызов основной функции
main();
