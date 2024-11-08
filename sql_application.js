// Подключаем библиотеку tedious
const { Connection } = require('tedious');

// Конфигурация подключения
const config = {  
    server: 'COMPUTER4630', // Имя сервера (или IP-адрес)
    authentication: {
        type: 'default',
        options: {
            userName: 'COMPUTER4630\\expert', // Имя пользователя SQL Server
            password: '' // Пароль пользователя SQL Server (оставьте пустым, если не используется)
        }
    },
    options: {
        encrypt: true, // Включите, если требуется шифрование (например, для Azure)
        trustServerCertificate: true, // Установите true, если используется самоподписанный сертификат
        database: 'Application_test', // Имя базы данных
        instanceName: 'SQLEXPRESS', // Имя экземпляра SQL Server
        connectTimeout: 15000, // Таймаут подключения в миллисекундах
        requestTimeout: 15000 // Таймаут запроса в миллисекундах
    }
};

// Создаем подключение
const connection = new Connection(config);

// Обработчик события подключения
connection.on('connect', (err) => {  
    if (err) {
        console.error("Ошибка подключения:", err.message);
    } else {
        console.log("Подключение установлено успешно!");
        
        // Закрываем соединение после успешного подключения
        connection.close();
    }
});

// Обработчик события ошибки (например, если соединение прервано)
connection.on('error', (err) => {
    console.error("Произошла ошибка:", err.message);
});

// Устанавливаем соединение
connection.connect();
