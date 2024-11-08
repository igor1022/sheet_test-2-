import express from 'express';
import { google } from 'googleapis';
import path from 'path';
import open from 'open'; // Используем import вместо require
import { writeFileSync } from 'fs';
import readlineSync from 'readline-sync';

const app = express();
const port = 3000;

const oAuth2Client = new google.auth.OAuth2(
  '385891869614-rb1225grm5vro613vsnvg56a0vtm500s.apps.googleusercontent.com',
  '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',
  'http://localhost:3000/oauth2callback'
);

const SCOPES = ['https://www.googleapis.com/auth/drive.readonly'];

app.get('/', (req, res) => {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  res.send(`<a href="${authUrl}">Авторизоваться через Google</a>`);
});

app.get('/oauth2callback', async (req, res) => {
  const { code } = req.query;
  
  if (!code) {
    return res.send('Ошибка: код авторизации не найден');
  }

  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    res.send('Авторизация прошла успешно!');
    
    writeFileSync('token.json', JSON.stringify(tokens));
    console.log('Токен сохранен в token.json');
  } catch (error) {
    res.send('Ошибка авторизации');
    console.error(error);
  }
});

app.listen(port, () => {
  console.log(`Сервер работает на http://localhost:${port}`);
  open(`http://localhost:${port}`);
});
