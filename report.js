import XLSX from 'xlsx';
import fs from 'fs';

// Используем встроенный fetch (Node.js 18+)
const FETCH_USER_APPS_ENDPOINT = 'https://backapi.rustore.ru/applicationData/retrieveUserApps?pagination=false';

// 🔐 НАСТРОЙКИ ТОКЕНОВ
const AUTH_TOKEN = process.env.RUSTORE_TOKEN;
const TELEGRAM_BOT_TOKEN = '8719192581:AAH8eQfyWHjZaLTvaGFeOQI-2LkGLLivNPk'; 
const TELEGRAM_CHAT_ID = '@rusotorestatsmy'; // Должен начинаться с -100

// Вспомогательная функция для паузы
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// 📅 ГЕНЕРАЦИЯ 4 ПЕРИОДОВ (по твоему фото)
const generateMonthPeriods = () => {
    const now = new Date();
    const year = now.getFullYear();
    const monthStr = String(now.getMonth() + 1).padStart(2, '0'); 
    const lastDay = new Date(year, now.getMonth() + 1, 0).getDate();

    return [
        { label: `01.${monthStr}-07.${monthStr}`, start: `${year}-${monthStr}-01`, end: `${year}-${monthStr}-07` },
        { label: `08.${monthStr}-14.${monthStr}`, start: `${year}-${monthStr}-08`, end: `${year}-${monthStr}-14` },
        { label: `15.${monthStr}-21.${monthStr}`, start: `${year}-${monthStr}-15`, end: `${year}-${monthStr}-21` },
        { label: `22.${monthStr}-${lastDay}.${monthStr}`, start: `${year}-${monthStr}-22`, end: `${year}-${monthStr}-${lastDay}` }
    ];
};

// 1️⃣ Получаем список всех игр
const getAllApps = async (attempts = 5) => {
    try {
        const response = await fetch(FETCH_USER_APPS_ENDPOINT, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': AUTH_TOKEN,
            },
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        return data.body.content || [];
    } catch (error) {
        if (attempts <= 0) throw new Error('Failed to fetch apps: ' + error.message);
        await delay(1000);
        return getAllApps(attempts - 1);
    }
};

// 2️⃣ Получаем доход за конкретный период
const getPeriodIncome = async (appId, dateFrom, dateTo, attempts = 3) => {
    const url = `https://api.rustore.ru/invoices-history/public/v1/apps/${appId}/invoice-payments/statistics?dateFrom=${dateFrom}T00:00:00Z&dateTo=${dateTo}T23:59:59Z`;
    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': AUTH_TOKEN,
            },
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        return data.body?.income?.sum?.overallSum?.totalStats || 0;
    } catch (error) {
        if (attempts <= 0) return 0;
        await delay(1000);
        return getPeriodIncome(appId, dateFrom, dateTo, attempts - 1);
    }
};

// 3️⃣ Отправка Excel в Telegram
const sendExcelToTelegram = async (filePath, fileName) => {
    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument`;
    
    const fileBuffer = fs.readFileSync(filePath);
    const fileBlob = new Blob([fileBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const formData = new FormData();
    formData.append('chat_id', TELEGRAM_CHAT_ID);
    formData.append('document', fileBlob, fileName);
    
    const monthName = new Date().toLocaleString('ru', { month: 'long', year: 'numeric' });
    formData.append('caption', `🟢 Отчет RuStore готов!\n📅 Период: ${monthName}`);

    try {
        const response = await fetch(url, { method: 'POST', body: formData });
        if (!response.ok) {
            const errDetails = await response.text();
            console.error('❌ Ошибка отправки в Telegram:', errDetails);
        } else {
            console.log('✅ Файл успешно отправлен!');
        }
    } catch (e) {
        console.error('❌ Ошибка сети:', e);
    }
};

// 🚀 ГЛАВНАЯ ФУНКЦИЯ
const runReport = async () => {
    console.log('🔄 Сбор данных запущен...');
    const periods = generateMonthPeriods();
    const apps = await getAllApps();
    
    const worksheetData = [
        ['Название игры', periods[0].label, periods[1].label, periods[2].label, periods[3].label, 'Итого за месяц']
    ];

    for (const app of apps) {
        // Защита названия от спецсимволов
        const safeAppName = app.appName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const rowData = [safeAppName];
        let appTotalIncome = 0;

        for (const period of periods) {
            const income = await getPeriodIncome(app.appId, period.start, period.end);
            rowData.push(income);
            appTotalIncome += income;
            await delay(500); 
        }

        // if (appTotalIncome === 0) continue; // Убери "//" в начале, если не хочешь видеть игры с нулевым доходом

        rowData.push(appTotalIncome);
        worksheetData.push(rowData);
    }

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Доходы');

    const fileName = 'RuStore_Monthly_Stats.xlsx';
    XLSX.writeFile(workbook, fileName);
    console.log('✅ Файл сгенерирован.');

    await sendExcelToTelegram(fileName, fileName);
};

runReport();
