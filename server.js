const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// КОНФИГУРАЦИЯ
const CONFIG = {
    currentRunUrl: 'https://api.outscore.ru/v1/partner/33/123/?key=cGwHeVOnIrz3va4StG6ReYG2BTBl1erZYhqVETASisw=&type=current',
    mainApiUrl: 'https://api.outscore.ru/v1/partner/33/123/?key=cGwHeVOnIrz3va4StG6ReYG2BTBl1erZYhqVETASisw=',
    excelPath: path.join(__dirname, 'data', 'broadcast_graphics.xlsx'),
    historyExcelPath: path.join(__dirname, 'data', 'attempts_history.xlsx'),
    backupDir: path.join(__dirname, 'logs'),
    maxBackups: 50
};

// Создаем директории
if (!fs.existsSync(path.join(__dirname, 'data'))) {
    fs.mkdirSync(path.join(__dirname, 'data'), { recursive: true });
}
if (!fs.existsSync(CONFIG.backupDir)) {
    fs.mkdirSync(CONFIG.backupDir, { recursive: true });
}

// Хранилище данных
let currentRunData = null;           // Текущий заезд
let completedRuns = [];               // Завершенные заезды (история)
let participantsCache = new Map();    // Кэш данных спортсменов (из Main API)
let resultsData = [];                 // Турнирная таблица
let competitionInfo = {};             // Информация о соревновании
let lastUpdateTime = null;

// ======================== НОРМАЛИЗАЦИЯ ========================

// Нормализация текущего заезда
function normalizeCurrentRun(rawCurrentRun) {
    if (!rawCurrentRun || !rawCurrentRun.user_id) return null;
    
    // Берем данные из кэша, если есть
    const cached = participantsCache.get(rawCurrentRun.user_id);
    
    return {
        id: rawCurrentRun.id,
        queuePos: rawCurrentRun.queue_pos,
        status: rawCurrentRun.status,
        stageId: rawCurrentRun.stage_id,
        stageName: rawCurrentRun.stage_id === 221 ? 'Квалификация' : 'Финал',
        groupId: rawCurrentRun.group_id,
        attempt: rawCurrentRun.attempt,
        attemptText: rawCurrentRun.attempt === 1 ? 'ПЕРВАЯ ПОПЫТКА' : (rawCurrentRun.attempt === 2 ? 'ВТОРАЯ ПОПЫТКА' : `${rawCurrentRun.attempt}-Я ПОПЫТКА`),
        userId: rawCurrentRun.user_id,
        riderName: cached?.fullName || `${rawCurrentRun.first_name || ''} ${rawCurrentRun.last_name || ''}`.trim(),
        riderFirstName: cached?.firstName || rawCurrentRun.first_name || '',
        riderLastName: cached?.lastName || rawCurrentRun.last_name || '',
        riderCity: cached?.city || '',
        riderRank: cached?.rank || '',
        riderIcon: cached?.icon || '',
        score: rawCurrentRun.score || rawCurrentRun.total || 0,
        flyTime: rawCurrentRun.fly_time || 0,
        tricks: rawCurrentRun.tricks || [],
        tricksTotal: rawCurrentRun.tricks_total || { TotalTricks: 0, TotalDegrees: 0, TotalTailWhips: 0, TotalBarSpins: 0 },
        timeStart: rawCurrentRun.time_start && rawCurrentRun.time_start !== '0001-01-01T00:00:00Z' ? rawCurrentRun.time_start : null,
        runTime: rawCurrentRun.run_time || 0
    };
}

// Обновление завершенных заездов
function updateCompletedRuns(currentRun) {
    if (!currentRun || currentRun.status !== 'finished') return;
    
    // Проверяем, есть ли уже этот заезд в истории
    const exists = completedRuns.some(r => r.id === currentRun.id);
    if (!exists && currentRun.score > 0) {
        completedRuns.unshift({
            id: currentRun.id,
            timestamp: new Date().toISOString(),
            stageId: currentRun.stageId,
            stageName: currentRun.stageName,
            attempt: currentRun.attempt,
            userId: currentRun.userId,
            riderName: currentRun.riderName,
            riderIcon: currentRun.riderIcon,
            score: currentRun.score,
            tricks: currentRun.tricks,
            tricksTotal: currentRun.tricksTotal
        });
        
        console.log(`✅ Новая попытка добавлена в историю: ${currentRun.riderName} - ${currentRun.score} баллов`);
        saveHistoryToExcel();
        updateResultsTable();
    }
}

// Обновление турнирной таблицы
function updateResultsTable() {
    // Группируем попытки по спортсменам
    const athleteMap = new Map();
    
    completedRuns.forEach(run => {
        if (!athleteMap.has(run.userId)) {
            athleteMap.set(run.userId, {
                userId: run.userId,
                riderName: run.riderName,
                riderIcon: run.riderIcon,
                attempts: {},
                bestScore: 0
            });
        }
        
        const athlete = athleteMap.get(run.userId);
        athlete.attempts[run.attempt] = run.score;
        
        // Обновляем лучший результат
        const currentBest = Math.max(...Object.values(athlete.attempts));
        athlete.bestScore = currentBest;
    });
    
    // Преобразуем в массив и сортируем
    resultsData = Array.from(athleteMap.values())
        .map((athlete, index) => ({
            place: index + 1,
            userId: athlete.userId,
            riderName: athlete.riderName,
            riderIcon: athlete.riderIcon,
            attempt1: athlete.attempts[1] || 0,
            attempt2: athlete.attempts[2] || 0,
            bestScore: athlete.bestScore
        }))
        .sort((a, b) => b.bestScore - a.bestScore)
        .map((item, idx) => ({ ...item, place: idx + 1 }));
    
    console.log(`📊 Турнирная таблица обновлена: ${resultsData.length} спортсменов`);
}

// ======================== ГРАФИЧЕСКИЕ ЛИСТЫ (СТРУКТУРА НЕ МЕНЯЕТСЯ) ========================

function createRiderProfileSheet(currentRun) {
    if (!currentRun) {
        return [['Статус'], ['Нет активного заезда']];
    }
    
    return [
        ['Имя спортсмена', 'Фото URL', 'Город', 'Разряд', 'Этап', 'Попытка'],
        [
            currentRun.riderName,
            currentRun.riderIcon,
            currentRun.riderCity,
            currentRun.riderRank,
            currentRun.stageName,
            currentRun.attemptText
        ]
    ];
}

function createFirstAttemptSheet() {
    const data = [['Место', 'Имя спортсмена', 'Фото URL', 'Город', 'Разряд', 'Баллы (1 попытка)']];
    
    const sortedResults = [...resultsData].sort((a, b) => a.place - b.place);
    
    sortedResults.forEach(result => {
        data.push([
            result.place,
            result.riderName,
            result.riderIcon,
            '', // Город - будет заполнен из Main API позже
            '', // Разряд - будет заполнен из Main API позже
            result.attempt1 > 0 ? result.attempt1.toFixed(2) : '0'
        ]);
    });
    
    return data;
}

function createLeaderboardSheet() {
    const data = [['Место', 'Имя спортсмена', 'Фото URL', 'Город', 'Разряд', 'Лучший результат', '1 попытка', '2 попытка']];
    
    const sortedResults = [...resultsData].sort((a, b) => b.bestScore - a.bestScore);
    
    sortedResults.forEach((result, idx) => {
        data.push([
            idx + 1,
            result.riderName,
            result.riderIcon,
            '', // Город
            '', // Разряд
            result.bestScore.toFixed(2),
            result.attempt1 > 0 ? result.attempt1.toFixed(2) : '-',
            result.attempt2 > 0 ? result.attempt2.toFixed(2) : '-'
        ]);
    });
    
    return data;
}

function createAthleteDetailedSheet() {
    const data = [['Место', 'Имя спортсмена', 'Фото URL', 'Город', 'Разряд', '1 попытка', '2 попытка', 'Лучший результат']];
    
    const sortedResults = [...resultsData].sort((a, b) => a.place - b.place);
    
    sortedResults.forEach(result => {
        data.push([
            result.place,
            result.riderName,
            result.riderIcon,
            '',
            '',
            result.attempt1 > 0 ? result.attempt1.toFixed(2) : '-',
            result.attempt2 > 0 ? result.attempt2.toFixed(2) : '-',
            result.bestScore.toFixed(2)
        ]);
    });
    
    return data;
}

function createCurrentRunGraphicsSheet(currentRun) {
    if (!currentRun) {
        return [['Статус'], ['Нет активного заезда']];
    }
    
    return [
        [
            'competition_name',
            'event_name',
            'rider_name',
            'rider_first_name',
            'rider_last_name',
            'rider_photo',
            'rider_city',
            'rider_rank',
            'stage_name',
            'attempt',
            'attempt_text',
            'current_score',
            'status',
            'tricks_count',
            'total_degrees',
            'total_tailwhips',
            'total_barspins'
        ],
        [
            competitionInfo.name || '',
            competitionInfo.eventName || '',
            currentRun.riderName,
            currentRun.riderFirstName,
            currentRun.riderLastName,
            currentRun.riderIcon,
            currentRun.riderCity,
            currentRun.riderRank,
            currentRun.stageName,
            currentRun.attempt,
            currentRun.attemptText,
            currentRun.score.toFixed(2),
            currentRun.status === 'finished' ? 'Завершено' : 'В процессе',
            currentRun.tricksTotal?.TotalTricks || 0,
            currentRun.tricksTotal?.TotalDegrees || 0,
            currentRun.tricksTotal?.TotalTailWhips || 0,
            currentRun.tricksTotal?.TotalBarSpins || 0
        ]
    ];
}

function createTricksGraphicsSheet(currentRun) {
    if (!currentRun || !currentRun.tricks || currentRun.tricks.length === 0) {
        return [['Статус'], ['Нет данных о трюках']];
    }
    
    const data = [['Время (сек)', 'Код трюка', 'Название трюка']];
    
    currentRun.tricks.forEach(trick => {
        data.push([
            trick.time?.toFixed(2) || '0',
            trick.code || '',
            trick.name || ''
        ]);
    });
    
    return data;
}

function createTop3Sheet() {
    const top3 = resultsData.filter(r => r.place <= 3);
    
    if (top3.length === 0) {
        return [['Статус'], ['Нет данных о призерах']];
    }
    
    const data = [['Место', 'Имя спортсмена', 'Фото URL', 'Город', 'Разряд', 'Лучший результат']];
    
    top3.forEach(result => {
        data.push([
            result.place,
            result.riderName,
            result.riderIcon,
            '',
            '',
            result.bestScore.toFixed(2)
        ]);
    });
    
    return data;
}

function createCompetitionSummarySheet() {
    return [
        ['Параметр', 'Значение'],
        ['Название соревнования', competitionInfo.name || 'Загрузка...'],
        ['Название этапа', competitionInfo.eventName || 'Загрузка...'],
        ['Даты проведения', competitionInfo.dates || '—'],
        ['Категория', competitionInfo.gender === 'M' ? 'Мужчины' : 'Женщины'],
        ['Всего участников', resultsData.length],
        ['Всего попыток', completedRuns.length],
        ['Последнее обновление', new Date().toLocaleString('ru-RU')]
    ];
}

// ======================== СОХРАНЕНИЕ В EXCEL ========================

async function saveGraphicsExcel() {
    try {
        const workbook = XLSX.utils.book_new();
        
        // Лист 1
        const riderProfileSheet = createRiderProfileSheet(currentRunData);
        const ws1 = XLSX.utils.aoa_to_sheet(riderProfileSheet);
        ws1['!cols'] = [{ wch: 20 }, { wch: 40 }, { wch: 20 }, { wch: 12 }, { wch: 15 }, { wch: 20 }];
        XLSX.utils.book_append_sheet(workbook, ws1, '1_Профиль_спортсмена');
        
        // Лист 2
        const firstAttemptSheet = createFirstAttemptSheet();
        const ws2 = XLSX.utils.aoa_to_sheet(firstAttemptSheet);
        ws2['!cols'] = [{ wch: 8 }, { wch: 25 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 18 }];
        XLSX.utils.book_append_sheet(workbook, ws2, '2_Первая_попытка');
        
        // Лист 3
        const leaderboardSheet = createLeaderboardSheet();
        const ws3 = XLSX.utils.aoa_to_sheet(leaderboardSheet);
        ws3['!cols'] = [{ wch: 8 }, { wch: 25 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 15 }, { wch: 12 }, { wch: 12 }];
        XLSX.utils.book_append_sheet(workbook, ws3, '3_Турнирная_таблица');
        
        // Лист 4
        const athleteDetailedSheet = createAthleteDetailedSheet();
        const ws4 = XLSX.utils.aoa_to_sheet(athleteDetailedSheet);
        ws4['!cols'] = [{ wch: 8 }, { wch: 25 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(workbook, ws4, '4_Детально_спортсмены');
        
        // Лист 5
        const currentRunGraphicsSheet = createCurrentRunGraphicsSheet(currentRunData);
        const ws5 = XLSX.utils.aoa_to_sheet(currentRunGraphicsSheet);
        XLSX.utils.book_append_sheet(workbook, ws5, '5_Текущий_заезд');
        
        // Лист 6
        const tricksSheet = createTricksGraphicsSheet(currentRunData);
        const ws6 = XLSX.utils.aoa_to_sheet(tricksSheet);
        ws6['!cols'] = [{ wch: 12 }, { wch: 15 }, { wch: 40 }];
        XLSX.utils.book_append_sheet(workbook, ws6, '6_Трюки');
        
        // Лист 7
        const top3Sheet = createTop3Sheet();
        const ws7 = XLSX.utils.aoa_to_sheet(top3Sheet);
        ws7['!cols'] = [{ wch: 8 }, { wch: 25 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(workbook, ws7, '7_Топ3');
        
        // Лист 8
        const summarySheet = createCompetitionSummarySheet();
        const ws8 = XLSX.utils.aoa_to_sheet(summarySheet);
        ws8['!cols'] = [{ wch: 25 }, { wch: 50 }];
        XLSX.utils.book_append_sheet(workbook, ws8, '8_Сводка');
        
        XLSX.writeFile(workbook, CONFIG.excelPath);
        console.log(`🎨 Excel для графики сохранен: ${CONFIG.excelPath}`);
        
        return true;
    } catch (error) {
        console.error('❌ Ошибка сохранения Excel:', error.message);
        return false;
    }
}

function saveHistoryToExcel() {
    try {
        if (completedRuns.length === 0) return;
        
        const historyData = completedRuns.map(r => ({
            'ID попытки': r.id,
            'Время': new Date(r.timestamp).toLocaleString('ru-RU'),
            'Этап': r.stageName,
            'Попытка': r.attempt,
            'ID спортсмена': r.userId,
            'Имя': r.riderName,
            'Фото': r.riderIcon,
            'Результат': r.score.toFixed(2),
            'Кол-во трюков': r.tricksTotal?.TotalTricks || 0,
            'Всего градусов': r.tricksTotal?.TotalDegrees || 0,
            'Tailwhips': r.tricksTotal?.TotalTailWhips || 0,
            'Barspins': r.tricksTotal?.TotalBarSpins || 0,
            'Список трюков': (r.tricks || []).map(t => t.name).join(' → ')
        }));
        
        const ws = XLSX.utils.json_to_sheet(historyData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'История_попыток');
        
        ws['!cols'] = [
            { wch: 12 }, { wch: 20 }, { wch: 15 }, { wch: 10 },
            { wch: 12 }, { wch: 25 }, { wch: 30 }, { wch: 12 },
            { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
            { wch: 50 }
        ];
        
        XLSX.writeFile(wb, CONFIG.historyExcelPath);
        console.log(`📜 История попыток сохранена: ${CONFIG.historyExcelPath} (${completedRuns.length} записей)`);
    } catch (error) {
        console.error('❌ Ошибка сохранения истории:', error.message);
    }
}

// ======================== ЗАГРУЗКА MAIN API (ОБОГАЩЕНИЕ ДАННЫХ) ========================

async function fetchMainData() {
    try {
        console.log('📡 Загрузка данных из Main API для обогащения...');
        const response = await axios.get(CONFIG.mainApiUrl, {
            timeout: 15000,
            headers: { 'Accept': 'application/json' }
        });
        
        const data = response.data;
        
        // Сохраняем информацию о соревновании
        if (data.competition) {
            competitionInfo.name = data.competition.name;
            competitionInfo.status = data.competition.status;
        }
        
        if (data.event) {
            competitionInfo.eventName = data.event.name;
            competitionInfo.dates = `${data.event.actual_start_date || ''} — ${data.event.actual_end_date || ''}`;
            competitionInfo.gender = data.event.gender;
        }
        
        // Кэшируем данные участников
        if (data.participants && Array.isArray(data.participants)) {
            data.participants.forEach(p => {
                participantsCache.set(p.id, {
                    id: p.id,
                    firstName: p.first_name || '',
                    lastName: p.last_name || '',
                    fullName: `${p.first_name || ''} ${p.last_name || ''}`.trim(),
                    city: p.city || '',
                    rank: p.rank || '',
                    icon: p.icon || ''
                });
            });
            console.log(`✅ Загружено ${participantsCache.size} участников из Main API`);
        }
        
        // Обрабатываем завершенные заезды из Main API (если есть)
        if (data.runs && Array.isArray(data.runs)) {
            data.runs.forEach(run => {
                if (run.status === 'finished' && run.score > 0) {
                    const cached = participantsCache.get(run.user_id);
                    const existingRun = completedRuns.some(r => r.id === run.id);
                    
                    if (!existingRun) {
                        completedRuns.push({
                            id: run.id,
                            timestamp: run.time_start || new Date().toISOString(),
                            stageId: run.stage_id,
                            stageName: run.stage_id === 221 ? 'Квалификация' : 'Финал',
                            attempt: run.attempt,
                            userId: run.user_id,
                            riderName: cached?.fullName || `${run.first_name || ''} ${run.last_name || ''}`.trim(),
                            riderIcon: cached?.icon || '',
                            score: run.score,
                            tricks: run.tricks || [],
                            tricksTotal: run.tricks_total || {}
                        });
                    }
                }
            });
            
            // Сортируем и обновляем
            completedRuns.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
            updateResultsTable();
            saveHistoryToExcel();
            console.log(`📊 Обработано ${data.runs.length} заездов из Main API`);
        }
        
        // Обновляем текущий заезд с enriched данными
        if (currentRunData) {
            const enriched = participantsCache.get(currentRunData.userId);
            if (enriched) {
                currentRunData.riderCity = enriched.city;
                currentRunData.riderRank = enriched.rank;
                currentRunData.riderIcon = enriched.icon;
                currentRunData.riderName = enriched.fullName;
            }
        }
        
        // Пересохраняем Excel с обогащенными данными
        await saveGraphicsExcel();
        
    } catch (error) {
        console.log(`⚠️ Main API недоступен: ${error.message}`);
        console.log('📌 Продолжаем работу только с Current API');
    }
}

// ======================== ОСНОВНАЯ ЛОГИКА (Current API как источник) ========================

let lastRunId = null;
let lastStatus = null;

async function fetchCurrentRun() {
    try {
        const response = await axios.get(CONFIG.currentRunUrl, {
            timeout: 10000,
            headers: { 'Accept': 'application/json' }
        });
        
        const rawData = response.data;
        const newRunData = normalizeCurrentRun(rawData);
        
        if (!newRunData) return;
        
        // Проверяем изменения
        const runChanged = lastRunId !== newRunData.id;
        const statusChanged = lastStatus !== newRunData.status;
        
        if (runChanged || statusChanged) {
            console.log(`🔄 Изменение: runId ${lastRunId} -> ${newRunData.id}, status ${lastStatus} -> ${newRunData.status}`);
            
            // Если предыдущий заезд завершился - сохраняем в историю
            if (lastRunId && lastRunId !== newRunData.id && lastStatus === 'finished') {
                // Находим завершенный заезд
                const completedRun = currentRunData;
                if (completedRun && completedRun.status === 'finished') {
                    updateCompletedRuns(completedRun);
                }
            }
        }
        
        // Обновляем текущие данные
        currentRunData = newRunData;
        lastRunId = currentRunData.id;
        lastStatus = currentRunData.status;
        
        // Обогащаем данными из кэша
        const enriched = participantsCache.get(currentRunData.userId);
        if (enriched) {
            currentRunData.riderCity = enriched.city;
            currentRunData.riderRank = enriched.rank;
            currentRunData.riderIcon = enriched.icon;
            currentRunData.riderName = enriched.fullName;
        }
        
        // Сохраняем в Excel
        await saveGraphicsExcel();
        
        console.log(`📡 Current API: ${currentRunData.riderName} - ${currentRunData.score} баллов (${currentRunData.status})`);
        
    } catch (error) {
        console.error(`❌ Ошибка Current API: ${error.message}`);
    }
}

// ======================== API ENDPOINTS ========================

// POST эндпоинт для приема данных (как вебхук)
app.post('/api/webhook/current', (req, res) => {
    const rawData = req.body;
    console.log('📨 Получен webhook с данными Current Run');
    
    const newRunData = normalizeCurrentRun(rawData);
    if (!newRunData) {
        return res.status(400).json({ error: 'Неверный формат данных' });
    }
    
    // Проверяем изменения
    if (lastRunId !== newRunData.id || lastStatus !== newRunData.status) {
        if (lastRunId && lastRunId !== newRunData.id && lastStatus === 'finished') {
            if (currentRunData && currentRunData.status === 'finished') {
                updateCompletedRuns(currentRunData);
            }
        }
    }
    
    currentRunData = newRunData;
    lastRunId = currentRunData.id;
    lastStatus = currentRunData.status;
    
    // Обогащаем
    const enriched = participantsCache.get(currentRunData.userId);
    if (enriched) {
        currentRunData.riderCity = enriched.city;
        currentRunData.riderRank = enriched.rank;
        currentRunData.riderIcon = enriched.icon;
        currentRunData.riderName = enriched.fullName;
    }
    
    saveGraphicsExcel();
    
    res.json({ 
        success: true, 
        message: 'Данные приняты',
        runId: currentRunData.id,
        status: currentRunData.status
    });
});

// GET эндпоинт для текущих данных
app.get('/api/current', (req, res) => {
    res.json({
        timestamp: new Date().toISOString(),
        currentRun: currentRunData,
        historyCount: completedRuns.length,
        participantsLoaded: participantsCache.size
    });
});

// GET эндпоинт для графики
app.get('/api/graphics', (req, res) => {
    res.json({
        timestamp: new Date().toISOString(),
        lastUpdate: lastUpdateTime,
        competition: competitionInfo,
        currentRun: currentRunData,
        leaderboard: resultsData,
        top3: resultsData.filter(r => r.place <= 3),
        historyCount: completedRuns.length
    });
});

// GET статистика
app.get('/api/stats', (req, res) => {
    res.json({
        lastUpdate: lastUpdateTime,
        currentRunExists: !!currentRunData,
        currentRunStatus: currentRunData?.status,
        historyRecordsCount: completedRuns.length,
        participantsInCache: participantsCache.size,
        resultsCount: resultsData.length,
        excelExists: fs.existsSync(CONFIG.excelPath),
        historyExcelExists: fs.existsSync(CONFIG.historyExcelPath)
    });
});

// GET скачать Excel
app.get('/api/download/graphics', (req, res) => {
    if (fs.existsSync(CONFIG.excelPath)) {
        res.download(CONFIG.excelPath, `broadcast_${new Date().toISOString().slice(0,19).replace(/:/g, '-')}.xlsx`);
    } else {
        res.status(404).json({ error: 'Файл еще не создан' });
    }
});

app.get('/api/download/history', (req, res) => {
    if (fs.existsSync(CONFIG.historyExcelPath)) {
        res.download(CONFIG.historyExcelPath, `attempts_history_${new Date().toISOString().slice(0,19).replace(/:/g, '-')}.xlsx`);
    } else {
        res.status(404).json({ error: 'История еще не создана' });
    }
});

// POST принудительное обновление
app.post('/api/refresh', async (req, res) => {
    await fetchCurrentRun();
    res.json({ success: true, message: 'Данные обновлены' });
});

// POST для загрузки Main API вручную
app.post('/api/load-main', async (req, res) => {
    await fetchMainData();
    res.json({ 
        success: true, 
        message: 'Main API загружен',
        participantsLoaded: participantsCache.size,
        historyRecords: completedRuns.length
    });
});

// ======================== ЗАПУСК ========================

// Запускаем polling Current API (каждые 10 секунд)
setInterval(fetchCurrentRun, 10000);

// Пытаемся загрузить Main API при старте и раз в минуту
setTimeout(() => {
    fetchMainData();
}, 2000);

setInterval(() => {
    fetchMainData();
}, 60000);

// Запускаем сервер
app.listen(PORT, () => {
    console.log(`
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║     BMX BROADCAST GRAPHICS SERVER v4.0                                      ║
    ║     🎨 Current API как основной источник                                   ║
    ╠══════════════════════════════════════════════════════════════════════════════╣
    ║  Сервер: http://localhost:${PORT}                                           ║
    ║                                                                              ║
    ║  📡 Источники данных:                                                       ║
    ║     ✅ Current API - основной (каждые 10 сек)                               ║
    ║     ⭕ Main API - дополнительный (обогащение фото/городов)                  ║
    ║                                                                              ║
    ║  🔗 API эндпоинты:                                                          ║
    ║     POST /api/webhook/current - вебхук для Current Run                      ║
    ║     GET  /api/current         - текущие данные                              ║
    ║     GET  /api/graphics        - данные для графики                          ║
    ║     GET  /api/stats           - статус сервера                              ║
    ║     POST /api/load-main       - принудительная загрузка Main API            ║
    ║     GET  /api/download/...    - скачать Excel                               ║
    ║                                                                              ║
    ║  💡 Система работает даже без Main API!                                     ║
    ║     (будут отсутствовать фото, города, разряды)                            ║
    ╚══════════════════════════════════════════════════════════════════════════════╝
    `);
});