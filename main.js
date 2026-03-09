import readline from 'readline/promises';
import { URL } from 'url';
import fs from 'fs';
import axios from 'axios';
import ExcelJS from 'exceljs';

// --- НАСТРОЙКИ ЗАДЕРЖЕК ---

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const getDelay = () => Math.random() * (5000 - 2000) + 2000; // 2 - 5 сек
const getDelaySafe = () => Math.random() * (10000 - 5000) + 5000; // 5 - 10 сек
const getDelayAggressive = () => Math.random() * (2000 - 1000) + 1000; // 1 - 2 сек

// --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

function checkString(s) {
    return /^(\d+%3B)*\d+$/.test(s);
}

function extractNumber(value) {
    if (typeof value !== 'string') return '';
    const match = value.match(/\d+(?:[.,]\d+)?/);
    if (match) {
        const numberStr = match[0].replace(',', '.');
        const parsed = parseFloat(numberStr);
        return isNaN(parsed) ? '' : parsed;
    }
    return '';
}

function safeGetField(obj, fieldName) {
    if (!obj || typeof obj !== 'object') return '';
    return obj[fieldName] !== undefined && obj[fieldName] !== null ? obj[fieldName] : '';
}

// --- ПАРСИНГ ВВОДА ---

function parseInput(inputStr) {
    const parts = inputStr.trim().split(/\s+/);

    if (parts.length > 2) {
        throw new Error("Необходимо указать максимум два параметра через пробел");
    }

    let sellerId = '';
    let brandId = '';

    if (parts.length === 2) {
        sellerId = parts[0];
        brandId = parts[1];

        if (!/^\d+$/.test(sellerId) || !checkString(brandId)) {
            throw new Error("Необходимо указать число (ID магазина) и ID бренда(ов)");
        }
    } else {
        const urlObj = new URL(inputStr);
        const pathParts = urlObj.pathname.split('/');
        
        // Ожидаем путь вида /seller/12345
        if (pathParts[1] === 'seller' && pathParts[2]) {
            sellerId = pathParts[2];
        } else {
            throw new Error("Не удалось извлечь ID продавца из ссылки");
        }

        brandId = urlObj.searchParams.get('fbrand');
        if (!brandId) {
            throw new Error("Не удалось извлечь fbrand из ссылки");
        }
    }

    return { sellerId, brandId };
}

// --- РАБОТА С API WB ---

async function getMediabasketRouteMap() {
    try {
        const url = 'https://cdn.wbbasket.ru/api/v3/upstreams';
        const response = await axios.get(url);
        
        if (response.data?.recommend?.mediabasket_route_map?.[0]?.hosts) {
            return response.data.recommend.mediabasket_route_map[0].hosts;
        }
        return [];
    } catch (e) {
        console.error(`Ошибка при получении route_map: ${e.message}`);
        return [];
    }
}

function getHostByRange(rangeValue, routeMap) {
    if (!routeMap || routeMap.length === 0) return '';

    for (const hostInfo of routeMap) {
        if (rangeValue >= hostInfo.vol_range_from && rangeValue <= hostInfo.vol_range_to) {
            return hostInfo.host;
        }
    }
    throw new Error(`Значение ${rangeValue} не попадает ни в один из доступных диапазонов корзин`);
}

async function fetchData(sellerId, brandId, routeMap) {
    const allProducts = [];
    const productsPerPage = 100;
    const urlTotalList = `https://catalog.wb.ru/sellers/v8/filters?ab_testing=false&appType=1&curr=rub&dest=12358357&fbrand=${brandId}&lang=ru&spp=30&supplier=${sellerId}&uclusters=0`;

    let productsTotal = 0;

    // Получаем общее количество товаров
    while (true) {
        try {
            const response = await axios.get(urlTotalList);
            productsTotal = response.data?.data?.total || 0;
            break;
        } catch (error) {
            if (error.response?.status === 429) {
                console.log('Слишком много запросов (filters). Безопасная задержка...');
                await delay(getDelaySafe());
            } else {
                throw error;
            }
        }
    }

    console.log(`Найдено товаров для парсинга: ${productsTotal}`);
    const pagesCount = Math.ceil(productsTotal / productsPerPage);
    let count = 1;

    for (let currentPage = 1; currentPage <= pagesCount; currentPage++) {
        const urlList = `https://catalog.wb.ru/sellers/v4/catalog?ab_testing=false&appType=1&curr=rub&dest=12358357&fbrand=${brandId}&hide_dtype=13&lang=ru&page=${currentPage}&sort=popular&spp=30&supplier=${sellerId}`;
        
        let products = [];
        while (true) {
            try {
                const response = await axios.get(urlList);
                products = response.data?.data?.products || response.data?.products || [];
                break;
            } catch (error) {
                if (error.response?.status === 429) {
                    console.log(`Слишком много запросов (catalog page ${currentPage}). Безопасная задержка...`);
                    await delay(getDelaySafe());
                } else {
                    console.error('Ошибка при получении списка товаров:', error.message);
                    await delay(getDelaySafe());
                }
            }
        }

        for (const item of products) {
            console.log(`Получено ${count}/${productsTotal}`);
            const productId = String(item.id);
            const vol = parseInt(productId.substring(0, productId.length - 5));
            const part = productId.substring(0, productId.length - 3);

            let basketName = getHostByRange(vol, routeMap);
            let basketNumber = 1;
            let isAutoServer = basketName.length > 0;
            let isNewLap = false;

            while (true) {
                const basketFormattedNumber = basketNumber < 10 ? `0${basketNumber}` : `${basketNumber}`;
                const host = isAutoServer ? basketName : `basket-${basketFormattedNumber}.wbbasket.ru`;
                const urlItem = `https://${host}/vol${vol}/part${part}/${productId}/info/ru/card.json`;

                try {
                    const productResponse = await axios.get(urlItem);
                    // Успех -> агрессивная (короткая) задержка
                    await delay(getDelayAggressive());
                    item.advanced = productResponse.data;
                    break;
                } catch (error) {
                    const status = error.response?.status;
                    if (status === 429) {
                        console.log('HTTP 429. Безопасная задержка...');
                        await delay(getDelay());
                    } else if (status === 404) {
                        // Перебор корзин, если сервер не найден
                        basketNumber = basketNumber < 24 ? basketNumber + 1 : 1;
                        if (isNewLap && basketNumber === 24) {
                            item.advanced = {}; // Если перебрали все 24 корзины и не нашли
                            break;
                        }
                        if (basketNumber === 1) isNewLap = true;
                    } else {
                        console.log(`Ошибка ${status || error.message}. Агрессивная задержка...`);
                        await delay(getDelaySafe());
                    }
                }
            }
            count++;
        }
        allProducts.push(...products);
    }
    return allProducts;
}

// --- МАППИНГ ДАННЫХ ---

const TARGET_COLUMNS = [
    'Группа', 'Артикул продавца', 'Артикул WB', 'Наименование', 'Категория продавца', 'Бренд', 'Описание', 'Фото', 'Видео', 'КИЗ',
    'Вес с упаковкой (кг)', '18+', 'Баркоды', 'Цена', 'Ставка НДС', 'Вес с упаковкой (кг)',
    'Высота упаковки', 'Длина упаковки', 'Ширина упаковки',
    'Дата окончания действия сертификата/декларации', 'Дата регистрации сертификата/декларации',
    'Номер декларации соответствия', 'Номер сертификата соответствия', 'Артикул OZON',
    'Вид автотранспорта', 'ИКПУ', 'Класс вязкости SAE', 'Классификация по ACEA',
    'Классификация по API', 'Классификация по CCMC', 'Классификация по ILSAC',
    'Классификация по ISO', 'Классификация по JASO', 'Классификация по NMMA',
    'Код упаковки', 'Комплектация', 'Назначение моторного масла', 'Совместимость',
    'Спецификация OEM', 'Срок годности', 'Страна производства', 'ТНВЭД',
    'Тип моторного масла', 'Упаковка', 'Артикул производителя', 'Модель',
    'ОЕМ номер', 'Объем (л)'
];

function mapData(data) {
    const newData = [];

    for (const item of data) {
        const advanced = item.advanced || {};
        let options = safeGetField(advanced, 'options');
        if (!Array.isArray(options)) options = [];

        let groupedOptions = safeGetField(advanced, 'grouped_options');
        if (!Array.isArray(groupedOptions)) groupedOptions = [];

        // Собираем все характеристики в один словарь (плоский список)
        const allCharacteristics = {};
        
        for (const opt of options) {
            if (opt?.name && opt?.value) {
                allCharacteristics[opt.name] = opt.value;
            }
        }

        for (const group of groupedOptions) {
            const groupOpts = group.options;
            if (Array.isArray(groupOpts)) {
                for (const opt of groupOpts) {
                    if (opt?.name && opt?.value) {
                        allCharacteristics[opt.name] = opt.value;
                    }
                }
            }
        }

        // WB хранит цену в копейках. Делим на 100
        const price = item.salePriceU ? item.salePriceU / 100 : '';

        const baseMappings = {
            'Группа': '',
            'Артикул продавца': item.supplierArticle || allCharacteristics['Артикул продавца'] || '',
            'Артикул WB': item.id || '',
            'Наименование': safeGetField(item, 'name'),
            'Категория продавца': safeGetField(item, 'entity'),
            'Бренд': safeGetField(item, 'brand'),
            'Описание': safeGetField(advanced, 'description'),
            'Фото': '',
            'Видео': '',
            'КИЗ': allCharacteristics['КИЗ'] || '',
            '18+': allCharacteristics['Ограничение'] || allCharacteristics['Возрастные ограничения'] || '',
            'Баркоды': '', 
            'Цена': price,
            'Ставка НДС': 20
        };

        const newItem = {};

        for (const col of TARGET_COLUMNS) {
            if (col in baseMappings) {
                newItem[col] = baseMappings[col];
            } else {
                let val = allCharacteristics[col] || '';

                // Очищаем габариты и вес от текстовых приписок, если они есть
                if (['Вес с упаковкой (кг)', 'Высота упаковки', 'Длина упаковки', 'Ширина упаковки'].includes(col) && val !== '') {
                    val = extractNumber(val);
                }

                newItem[col] = val;
            }
        }

        newData.push(newItem);
    }

    return newData;
}

// --- EXCEL ---

function generateFilename(baseName = "result") {
    const now = new Date();
    const pad = (n) => n.toString().padStart(2, '0');
    
    const year = now.getFullYear();
    const month = pad(now.getMonth() + 1);
    const day = pad(now.getDate());
    const hours = pad(now.getHours());
    const minutes = pad(now.getMinutes());
    const seconds = pad(now.getSeconds());

    return `${baseName}_${year}-${month}-${day}_${hours}-${minutes}-${seconds}.xlsx`;
}

async function createFile(data) {
    const templatePath = 'wb_template.xlsx';
    const outputPath = generateFilename();

    if (!fs.existsSync(templatePath)) {
        console.error(`Ошибка: Файл шаблона '${templatePath}' не найден в папке скрипта!`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    
    const worksheet = workbook.worksheets[0]; // Берем первый лист
    
    // Запись спарсенных товаров начинается строго с 5-й строки
    let startRow = 5; 

    for (const item of data) {
        const row = worksheet.getRow(startRow);
        
        // Записываем данные по порядку целевых колонок
        TARGET_COLUMNS.forEach((colName, index) => {
            // Excel-колонки начинаются с 1 (поэтому index + 1)
            row.getCell(index + 1).value = item[colName];
        });
        
        row.commit();
        startRow++;
    }

    await workbook.xlsx.writeFile(outputPath);
    console.log(`Данные успешно записаны в файл: ${outputPath}`);
}

// --- ОСНОВНАЯ ФУНКЦИЯ ---

async function main() {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    try {
        const answer = await rl.question('Введите ссылку WB или ID магазина и ID бренда(ов) через пробел: ');
        rl.close();

        console.log('Парсинг ввода...');
        const { sellerId, brandId } = parseInput(answer);
        console.log(`Seller ID: ${sellerId}, Brand ID: ${brandId}`);

        console.log('Получение распределения серверов (маршрутизация корзин)...');
        const routeMap = await getMediabasketRouteMap();

        console.log('Получение списка товаров и их характеристик...');
        const rawData = await fetchData(sellerId, brandId, routeMap);

        console.log('Обработка и маппинг характеристик...');
        const mappedData = mapData(rawData);

        console.log('Создание Excel файла...');
        await createFile(mappedData);

        console.log('Готово!');
    } catch (error) {
        console.error('Критическая ошибка:', error.message);
        rl.close();
    }
}

// Запуск скрипта
main();
