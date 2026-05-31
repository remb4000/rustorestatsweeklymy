import XLSX from 'xlsx';
import fs from 'fs';

const FETCH_USER_APPS_ENDPOINT = 'https://backapi.rustore.ru/applicationData/retrieveUserApps?pagination=false';

// 🔐 НАСТРОЙКИ
const AUTH_TOKEN = process.env.RUSTORE_TOKEN ;
const TELEGRAM_BOT_TOKEN = '8719192581:AAH8eQfyWHjZaLTvaGFeOQI-2LkGLLivNPk'; 
const TELEGRAM_CHAT_ID = '@rusotorestatsmy'; 

// ⏳ УМНЫЕ ПАУЗЫ
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
const randomDelay = (min, max) => new Promise(resolve => setTimeout(resolve, Math.floor(Math.random() * (max - min + 1)) + min));

const now = new Date();
const CURRENT_YEAR = now.getFullYear();
const CURRENT_MONTH = now.getMonth() + 1; 
const lastDay = new Date(CURRENT_YEAR, CURRENT_MONTH, 0).getDate();
const monthStr = String(CURRENT_MONTH).padStart(2, '0');

// 🔥 Точка отсечения (Первое число текущего месяца)
// Чеки старше этой даты мы качать не будем!
const startOfCurrentMonth = new Date(CURRENT_YEAR, CURRENT_MONTH - 1, 1);

const PERIOD_LABELS = [
    `01.${monthStr}-07.${monthStr}`,
    `08.${monthStr}-14.${monthStr}`,
    `15.${monthStr}-21.${monthStr}`,
    `22.${monthStr}-${lastDay}.${monthStr}`
];

// 1️⃣ Получаем список игр
const getAllApps = async (attempts = 3) => {
    try {
        const response = await fetch(FETCH_USER_APPS_ENDPOINT, {
            headers: { 'Content-Type': 'application/json', 'Authorization': AUTH_TOKEN },
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const data = await response.json();
        return data.body.content || [];
    } catch (error) {
        if (attempts <= 0) return [];
        await delay(2000);
        return getAllApps(attempts - 1);
    }
};

// 2️⃣ СКАЧИВАНИЕ ЧЕКОВ С ОСТАНОВКОЙ НА ПРОШЛОМ МЕСЯЦЕ
const getAppInvoices = async (appId, appName, attempts = 5) => {
    let allInvoices = [];
    let page = 0;
    let totalPages = 1;

    while (page < totalPages) {
        const url = `https://api.rustore.ru/v1/monetization/invoices-history/apps/${appId}/invoice-payments?page=${page}&size=100`;
        
        try {
            const response = await fetch(url, {
                headers: { 'Content-Type': 'application/json', 'Authorization': AUTH_TOKEN },
            });
            
            // 🛑 ЖЕСТКАЯ ПАУЗА ПРИ 429
            if (response.status === 429) {
                console.warn(`\n🛑 RuStore ругается (Код 429). Уходим в спячку на 35 секунд...`);
                await delay(35000); // 35 секунд штрафа (наверняка)
                throw new Error('Rate Limited');
            }

            if (!response.ok) {
                const errText = await response.text();
                console.error(`\n❌ ОШИБКА СЕРВЕРА ДЛЯ "${appName}": Код ${response.status}`);
                throw new Error(`HTTP ${response.status}`);
            }

            const data = await response.json();
            
            if (data.invoices && data.invoices.length > 0) {
                allInvoices = allInvoices.concat(data.invoices);
                
                // 🧠 УМНЫЙ ТОРМОЗ: Проверяем дату ПОСЛЕДНЕГО чека на странице
                const oldestInvoiceOnPage = new Date(data.invoices[data.invoices.length - 1].invoice_date);
                if (oldestInvoiceOnPage < startOfCurrentMonth) {
                    console.log(`   ⏭️ Дошли до старых чеков (${oldestInvoiceOnPage.toLocaleDateString()}). Останавливаем загрузку истории.`);
                    break; // Выходим из цикла while, дальше качать страницы не нужно!
                }
            }
            
            totalPages = data.totalPages || 1;
            page++;
            
            // Пауза между страницами (если у игры много транзакций в ЭТОМ месяце)
            if (page < totalPages) {
                await randomDelay(1500, 2500); 
            }
            
        } catch (error) {
            if (attempts <= 1) return []; 
            console.log(`⚠️ Повторная попытка скачивания... (осталось попыток: ${attempts - 1})`);
            return getAppInvoices(appId, appName, attempts - 1);
        }
    }
    return allInvoices;
};

// 3️⃣ Отправка Excel в Telegram
const sendExcelToTelegram = async (filePath, fileName) => {
    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument`;
    const fileBuffer = fs.readFileSync(filePath);
    const fileBlob = new Blob([fileBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const formData = new FormData();
    formData.append('chat_id', TELEGRAM_CHAT_ID);
    formData.append('document', fileBlob, fileName);
    formData.append('caption', `🟢 Отчет по транзакциям RuStore готов!\n📅 Период: ${now.toLocaleString('ru', { month: 'long', year: 'numeric' })}`);

    try {
        const response = await fetch(url, { method: 'POST', body: formData });
        if (response.ok) console.log('\n✅ Файл успешно отправлен в Telegram!');
        else console.error('❌ Ошибка отправки в ТГ:', await response.text());
    } catch (e) {
        console.error('❌ Ошибка сети:', e);
    }
};

// 🚀 ГЛАВНАЯ ФУНКЦИЯ
const runReport = async () => {
    console.log('🔄 Начинаем быстрый и безопасный сбор чеков...');
    const apps = await getAllApps();
    
    if (apps.length === 0) {
        console.log('❌ Не удалось получить список игр.');
        return;
    }

    const worksheetData = [
        ['Название игры', PERIOD_LABELS[0], PERIOD_LABELS[1], PERIOD_LABELS[2], PERIOD_LABELS[3], 'Итого за месяц']
    ];

    for (const app of apps) {
        console.log(`\n📊 Игра: ${app.appName}`);
        const safeAppName = app.appName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        
        let incomeByWeek = [0, 0, 0, 0];
        let totalIncome = 0;

        const invoices = await getAppInvoices(app.appId, app.appName);
        
        let validThisMonth = 0;

        for (const invoice of invoices) {
            if (invoice.invoice_status !== 'confirmed' && invoice.invoice_status !== 'paid') continue;

            const date = new Date(invoice.invoice_date);
            if (date.getFullYear() !== CURRENT_YEAR || (date.getMonth() + 1) !== CURRENT_MONTH) continue;

            const amount = invoice.amount_create / 100;
            const day = date.getDate();

            if (day >= 1 && day <= 7) incomeByWeek[0] += amount;
            else if (day >= 8 && day <= 14) incomeByWeek[1] += amount;
            else if (day >= 15 && day <= 21) incomeByWeek[2] += amount;
            else if (day >= 22) incomeByWeek[3] += amount;
            
            totalIncome += amount;
            validThisMonth++;
        }

        console.log(`   ✅ Обработано успешных чеков за этот месяц: ${validThisMonth}`);
        if (validThisMonth > 0) {
            console.log(`   💰 Сумма за месяц: ${totalIncome.toFixed(2)} ₽`);
        }

        worksheetData.push([
            safeAppName, 
            parseFloat(incomeByWeek[0].toFixed(2)), 
            parseFloat(incomeByWeek[1].toFixed(2)), 
            parseFloat(incomeByWeek[2].toFixed(2)), 
            parseFloat(incomeByWeek[3].toFixed(2)), 
            parseFloat(totalIncome.toFixed(2))
        ]);
        
        // Пауза перед следующей игрой
        await randomDelay(2000, 4000); 
    }

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Доходы');

    const fileName = 'RuStore_Invoices_Stats.xlsx';
    XLSX.writeFile(workbook, fileName);
    console.log('\n✅ Excel файл успешно сформирован.');

    await sendExcelToTelegram(fileName, fileName);
};

runReport();
