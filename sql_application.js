const sql = require('mssql');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Конфигурация для MSSQL подключения
const config = {
    user: 'dbadmin',
    password: 'Account8',
    server: 'sheetstest.database.windows.net',
    database: 'testdb', // Указываем нужную базу данных
    options: {
        encrypt: true,
        trustServerCertificate: false
    }
};

// Функция для создания таблицы, если она еще не существует
async function createTableIfNotExists() {
    try {
        // Подключение к базе данных
        const pool = await sql.connect(config);

        // SQL запрос для создания таблицы, если она не существует
        const createTableQuery = `
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'transfer_applications')
            BEGIN
                CREATE TABLE transfer_applications (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    timestamp DATETIME,
                    surname NVARCHAR(255),
                    clientType NVARCHAR(255),
                    direction NVARCHAR(255),
                    clientRequest NVARCHAR(255),
                    presentationLink NVARCHAR(255),
                    phoneNumber NVARCHAR(255),
                    clientName NVARCHAR(255),
                    budget FLOAT,
                    reasonForNonTarget NVARCHAR(255),
                    requestLink NVARCHAR(255),
                    adType NVARCHAR(255),
                    department NVARCHAR(255),
                    source NVARCHAR(255)
                );
            END
        `;

        // Выполнение запроса
        await pool.request().query(createTableQuery);
        console.log('Таблица transfer_applications успешно создана или уже существует.');
    } catch (err) {
        console.error('Ошибка при создании таблицы:', err);
    }
}

// Функция для вставки данных в базу данных
async function insertData(data) {
    try {
        // Подключение к базе данных
        await sql.connect(config);
        console.log("Подключение к базе данных успешно.");

        // Перебор массива данных и вставка каждого элемента в базу данных
        for (let row of data) {
            const request = new sql.Request();
            request.input('timestamp', sql.DateTime, row["Отметка времени"]);
            request.input('surname', sql.NVarChar, row["Фамилия"]);
            request.input('clientType', sql.NVarChar, row["Тип клиента"]);
            request.input('direction', sql.NVarChar, row["Направление"]);
            request.input('clientRequest', sql.NVarChar, row["Запрос клиента"]);
            request.input('presentationLink', sql.NVarChar, row["Ссылка на презентацию"]);
            request.input('phoneNumber', sql.NVarChar, processPhoneNumber(row["Номер телефона клиента (380..)"])); // Обрабатываем телефонный номер
            request.input('clientName', sql.NVarChar, row["Имя клиента"]);

            // Проверка и обработка значения для 'budget'
            const budgetValue = parseFloat(row["Бюджет"]);
            request.input('budget', sql.Float, isNaN(budgetValue) ? null : budgetValue);

            request.input('reasonForNonTarget', sql.NVarChar, row["Причина нецевого клиента"]);
            request.input('requestLink', sql.NVarChar, row["Ссылка на заявку"]);
            request.input('adType', sql.NVarChar, row["Тип рекламы"]);
            request.input('department', sql.NVarChar, row["Передано в отдел"]);
            request.input('source', sql.NVarChar, row["Источник"]);

            const query = `
                INSERT INTO transfer_applications (
                    timestamp,
                    surname,
                    clientType,
                    direction,
                    clientRequest,
                    presentationLink,
                    phoneNumber,
                    clientName,
                    budget,
                    reasonForNonTarget,
                    requestLink,
                    adType,
                    department,
                    source
                ) VALUES (
                    @timestamp,
                    @surname,
                    @clientType,
                    @direction,
                    @clientRequest,
                    @presentationLink,
                    @phoneNumber,
                    @clientName,
                    @budget,
                    @reasonForNonTarget,
                    @requestLink,
                    @adType,
                    @department,
                    @source
                )`;

            // Выполнение запроса
            await request.query(query);
        }

        console.log("Данные успешно вставлены.");
    } catch (err) {
        console.error("Ошибка при подключении к базе данных или импорте данных:", err);
    } finally {
        // Закрытие подключения к базе данных
        sql.close();
    }
}

// Функция для обработки номера телефона
function processPhoneNumber(phoneNumber) {
    if (!phoneNumber) {
        return null;
    }

    // Удаляем все символы, кроме цифр
    const cleanedPhoneNumber = phoneNumber.replace(/\D/g, "");

    // Если номер начинается с 380, оставляем его как есть
    if (cleanedPhoneNumber.startsWith("380")) {
        return cleanedPhoneNumber;
    }

    // Если это не номер в нужном формате, возвращаем null
    return null;
}

// Функция для загрузки данных из Excel файла
function loadDataFromExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Берем первый лист
    const worksheet = workbook.Sheets[sheetName];

    // Преобразуем данные листа в JSON
    const data = XLSX.utils.sheet_to_json(worksheet);
    console.log("Данные из Excel файла успешно прочитаны:");
    console.log(data); // Выводим все данные для диагностики

    // Вставляем данные в базу данных
    insertData(data);
}

// Укажите путь к вашему файлу Excel
const filePath = "output.xlsx";

// Запуск создания таблицы и загрузки данных
createTableIfNotExists().then(() => {
    loadDataFromExcel(filePath);
});
