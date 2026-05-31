import XLSX from 'xlsx';
import fs from 'fs';

const FETCH_USER_APPS_ENDPOINT = 'https://backapi.rustore.ru/applicationData/retrieveUserApps?pagination=false';

// 🔐 НАСТРОЙКИ
const AUTH_TOKEN = process.env.RUSTORE_TOKEN;
const TELEGRAM_BOT_TOKEN = '8719192581:AAH8eQfyWHjZaLTvaGFeOQI-2LkGLLivNPk'; 
const TELEGRAM_CHAT_ID = '@rusotorestatsmy'; // С минусом, например -1004234492621

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// 📅 Получаем текущий месяц для фильтрации
const now = new Date();
const CURRENT_YEAR = now.getFullYear();
const CURRENT_MONTH = now.getMonth() + 1; // getMonth отдает 0-11
const lastDay = new Date(CURRENT_YEAR, CURRENT_MONTH, 0).getDate();
const monthStr = String(CURRENT_MONTH).padStart(2, '0');

// Названия колонок для Excel
const PERIOD_LABELS = [
    `01.${monthStr}-07.${monthStr}`,
    `08.${monthStr}-14.${monthStr}`,
    `15.${monthStr}-21.${monthStr}`,
    `22.${monthStr}-${lastDay}.${monthStr}`
];

// 1️⃣ Получаем список всех игр
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
        await delay(1000);
        return getAllApps(attempts - 1);
    }
};

// 2️⃣ НОВАЯ ЛОГИКА: Скачиваем все чеки игры (по 100 за раз)
const getAppInvoices = async (appId, attempts = 3) => {
    let allInvoices = [];
    let page = 0;
    let totalPages = 1;

    while (page < totalPages) {
        // Берем сразу по 100 чеков, чтобы не делать много запросов
        const url = `https://api.rustore.ru/v1/monetization/invoices-history/apps/${appId}/invoice-payments?page=${page}&size=100&invoiceStatuses=CONFIRMED&invoiceStatuses=PAID`;
        
        try {
            const response = await fetch(url, {
                headers: { 'Content-Type': 'application/json', 'Authorization': AUTH_TOKEN },
            });
            
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const data = await response.json();
            
            if (data.invoices && data.invoices.length > 0) {
                allInvoices = allInvoices.concat(data.invoices);
            }
            
            totalPages = data.totalPages || 1;
            page++;
            await delay(300); // Пауза между страницами
        } catch (error) {
            if (attempts <= 0) break;
            await delay(1000);
            return getAppInvoices(appId, attempts - 1);
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
        if (response.ok) console.log('✅ Файл успешно отправлен!');
        else console.error('❌ Ошибка отправки:', await response.text());
    } catch (e) {
        console.error('❌ Ошибка сети:', e);
    }
};

// 🚀 ГЛАВНАЯ ФУНКЦИЯ
const runReport = async () => {
    console.log('🔄 Начинаем скачивание чеков...');
    const apps = await getAllApps();
    
    // Шапка Excel
    const worksheetData = [
        ['Название игры', PERIOD_LABELS[0], PERIOD_LABELS[1], PERIOD_LABELS[2], PERIOD_LABELS[3], 'Итого за месяц']
    ];

    for (const app of apps) {
        console.log(`📊 Обработка игры: ${app.appName}`);
        const safeAppName = app.appName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        
        // Массив для хранения сумм по неделям (индексы 0-3) и итого (4)
        let incomeByWeek = [0, 0, 0, 0];
        let totalIncome = 0;

        // Скачиваем чеки
        const invoices = await getAppInvoices(app.appId);

        // Раскидываем чеки по неделям
        for (const invoice of invoices) {
            // Пропускаем отмененные и возвраты (на всякий случай, хоть мы и фильтровали в URL)
            if (invoice.invoice_status !== 'confirmed' && invoice.invoice_status !== 'paid') continue;

            const date = new Date(invoice.invoice_date);
            
            // Если чек не из этого месяца/года - пропускаем
            if (date.getFullYear() !== CURRENT_YEAR || (date.getMonth() + 1) !== CURRENT_MONTH) continue;

            // Сумма в API лежит в копейках (19999), делим на 100 -> 199.99
            const amount = invoice.amount_create / 100;
            const day = date.getDate(); // Получаем число (от 1 до 31)

            // Сортируем по корзинам
            if (day >= 1 && day <= 7) {
                incomeByWeek[0] += amount;
            } else if (day >= 8 && day <= 14) {
                incomeByWeek[1] += amount;
            } else if (day >= 15 && day <= 21) {
                incomeByWeek[2] += amount;
            } else if (day >= 22) {
                incomeByWeek[3] += amount;
            }
            
            totalIncome += amount;
        }

        // Если игра не принесла доход в этом месяце - не выводим ее в Excel (по желанию можно убрать эту строчку)
        // if (totalIncome === 0) continue; 

        // Округляем до 2 знаков после запятой, чтобы не было длинных хвостов вроде 199.989999
        worksheetData.push([
            safeAppName, 
            parseFloat(incomeByWeek[0].toFixed(2)), 
            parseFloat(incomeByWeek[1].toFixed(2)), 
            parseFloat(incomeByWeek[2].toFixed(2)), 
            parseFloat(incomeByWeek[3].toFixed(2)), 
            parseFloat(totalIncome.toFixed(2))
        ]);
        
        await delay(500); 
    }

    // 💾 Создаем Excel
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Доходы');

    const fileName = 'RuStore_Invoices_Stats.xlsx';
    XLSX.writeFile(workbook, fileName);
    console.log('✅ Excel файл успешно сформирован по реальным чекам.');

    await sendExcelToTelegram(fileName, fileName);
};

runReport();
