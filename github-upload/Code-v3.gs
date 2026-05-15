// ============================================================
// СТО УЧЁТ — Google Apps Script v1.2
// Поддержка JSONP для работы с GitHub Pages
// ============================================================

const SHEET_RECORDS = 'Записи';
const SHEET_REFS = 'Справочники';
const SHEET_USERS = 'Пользователи';
const SHEET_NOMENCLATURE = 'Номенклатура';
const SHEET_WAREHOUSE = 'Склад';

function doGet(e) {
  try {
    const p = e.parameter;
    const callback = p.callback || null;
    let result;

    if (p._method === 'POST' && p.action === 'saveRecord') {
      const rows = JSON.parse(decodeURIComponent(p.rows || '[]'));
      result = saveRecord(rows);
    } else if (p.action === 'getRefs') {
      result = getRefs();
    } else if (p.action === 'getUser') {
      result = getUser(p.email);
    } else if (p.action === 'getNomenclature') {
      result = getNomenclature();
    } else if (p.action === 'saveNomenclature') {
      const record = JSON.parse(decodeURIComponent(p.record || '{}'));
      result = saveNomenclature(record);
    } else if (p.action === 'saveWarehouse') {
      const record = JSON.parse(decodeURIComponent(p.record || '{}'));
      result = saveWarehouse(record);
    } else {
      result = { error: 'Неизвестное действие: ' + p.action };
    }

    return buildResponse(result, callback);
  } catch (err) {
    return buildResponse({ error: err.toString() }, e.parameter?.callback);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    if (data.action === 'saveRecord') {
      result = saveRecord(data.rows);
    } else if (data.action === 'saveNomenclature') {
      result = saveNomenclature(data.record);
    } else if (data.action === 'saveWarehouse') {
      result = saveWarehouse(data.record);
    } else {
      result = { error: 'Неизвестное действие' };
    }

    return buildResponse(result, null);
  } catch (err) {
    return buildResponse({ error: err.toString() }, null);
  }
}

function saveRecord(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RECORDS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RECORDS);
    const headers = ['ID','Дата','Исполнитель','Email','Группа услуг',
                     'Марка/Модель','Номер авто','VIN','Косяк',
                     'Услуга/Работа','Комментарий','Материал','Количество','Ед.изм','NR'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1a1a2e').setFontColor('#f0a500').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  rows.forEach(row => {
    sheet.appendRow([
      row.ID, row.Дата, row.Исполнитель, row.Email, row.Группа_услуг,
      row.Марка_модель, row.Номер_авто, row.VIN, row.Косяк, row.Услуга,
      row.Комментарий, row.Материал, row.Количество, row.Ед_изм, row.NR
    ]);
  });

  return { success: true, saved: rows.length };
}

function getUser(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);

  if (!sheet) {
    createUsersTemplate();
    return { error: 'Лист Пользователи создан — заполните его.' };
  }

  const data = sheet.getDataRange().getValues();
  const searchEmail = String(email || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const userEmail = String(data[i][3] || '').trim().toLowerCase();
    if (userEmail === searchEmail) {
      return {
        success: true,
        user: {
          name: String(data[i][0] || '').trim(),
          role: String(data[i][1] || '').trim(),
          group: String(data[i][2] || '').trim(),
          email: userEmail
        }
      };
    }
  }

  return { error: 'Пользователь не найден' };
}

function getRefs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REFS);

  if (!sheet) {
    createRefsTemplate();
    return { error: 'Лист Справочники создан — заполните его.' };
  }

  const data = sheet.getDataRange().getValues();
  const refs = { services: {}, materials: {}, cars: [] };

  for (let i = 1; i < data.length; i++) {
    const group = String(data[i][0] || '').trim();
    const service = String(data[i][1] || '').trim();
    const material = String(data[i][2] || '').trim();
    const unit = String(data[i][3] || '').trim();

    if (!group || !service) continue;
    if (!refs.services[group]) refs.services[group] = [];
    if (!refs.services[group].includes(service)) refs.services[group].push(service);
    if (material) {
      if (!refs.materials[service]) refs.materials[service] = [];
      if (!refs.materials[service].find(m => m.name === material)) {
        refs.materials[service].push({ name: material, unit });
      }
    }
  }

  const carsSheet = ss.getSheetByName('Авто');
  if (carsSheet) {
    carsSheet.getDataRange().getValues().slice(1).forEach(r => {
      const model = String(r[0] || '').trim();
      const plate = String(r[1] || '').trim();
      const vin = String(r[2] || '').trim();
      if (model) refs.cars.push({ model, plate, vin });
    });
  }

  return { success: true, refs };
}

function getNomenclature() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NOMENКЛАТУРА);

  if (!sheet) {
    createNomenclatureTemplate();
    return { error: 'Лист Номенклатура создан — заполните его.' };
  }

  const data = sheet.getDataRange().getValues();
  const nomenclature = [];
  for (let i = 1; i < data.length; i++) {
    const barcode = String(data[i][0] || '').trim();
    const name = String(data[i][1] || '').trim();
    const unit = String(data[i][2] || '').trim();
    const comment = String(data[i][3] || '').trim();
    const registrar = String(data[i][4] || '').trim();
    const date = String(data[i][5] || '').trim();
    if (barcode && name) nomenclature.push({ barcode, name, unit, comment, registrar, date });
  }
  return { success: true, nomenclature };
}

function saveNomenclature(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NOMENКЛАТУРА);
  if (!sheet) {
    sheet = createNomenclatureTemplate();
  }
  sheet.appendRow([
    record.Штрихкод || '',
    record.Название || '',
    record.Ед_изм || '',
    record.Комментарий || '',
    record.Регистратор || '',
    record.Дата || ''
  ]);
  return { success: true };
}

function saveWarehouse(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_WAREHOUSE);
  if (!sheet) {
    sheet = createWarehouseTemplate();
  }
  sheet.appendRow([
    record.ID || '',
    record.Дата || '',
    record.Штрихкод || '',
    record.Название || '',
    record['Ед. изм.'] || '',
    record.Комментарий || '',
    record.Ответственный || '',
    record.Статус || '',
    record.Количество || ''
  ]);
  return { success: true };
}

function createNomenclatureTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_NOMENКЛАТУРА);
  sheet.getRange(1, 1, 1, 6).setValues([['Штрихкод','Название','Ед. изм.','Комментарий','Регистратор','Дата']])
    .setBackground('#1a1a2e').setFontColor('#f0a500').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function createWarehouseTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_WAREHOUSE);
  sheet.getRange(1, 1, 1, 9).setValues([['ID','Дата','Штрихкод','Название','Ед. изм.','Комментарий','Ответственный','Статус','Количество']])
    .setBackground('#1a1a2e').setFontColor('#f0a500').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function createRefsTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_REFS);
  sheet.getRange(1, 1, 1, 4).setValues([['Группа услуг','Услуга / Работа','Материал','Ед. изм.']])
    .setBackground('#1a1a2e').setFontColor('#f0a500').setFontWeight('bold');
  const ex = [
    ['Мастерская','Замена масла и фильтра','Масло моторное 5W40','л'],
    ['Мастерская','Замена масла и фильтра','Фильтр масляный','шт'],
    ['Склад','Выдача расходников','Ветошь','кг'],
    ['Плоттер','Резка виниловой плёнки','Плёнка матовая','м²'],
  ];
  sheet.getRange(2, 1, ex.length, 4).setValues(ex);
  sheet.setFrozenRows(1);
}

function createUsersTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_USERS);
  sheet.getRange(1, 1, 1, 4).setValues([['ФИО','Роль','Группа услуг','Gmail']])
    .setBackground('#1a1a2e').setFontColor('#f0a500').setFontWeight('bold');
  sheet.getRange(2, 1, 3, 4).setValues([
    ['Иван Петров','Мастер','Мастерская','ivan@gmail.com'],
    ['Сергей Ковалёв','Кладовщик','Склад','sergey@gmail.com'],
    ['Анна Мищенко','Плоттер','Плоттер','anna@gmail.com'],
  ]);
  sheet.setFrozenRows(1);
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function buildResponse(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
