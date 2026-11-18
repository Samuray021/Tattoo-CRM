const rubles = {type: 'CURRENCY', pattern: '#,##0.00 "₽"'};
const dateFormat = {type: 'DATE', pattern: 'dd.mm.yyyy'};
const monthFormat = {type: 'DATE', pattern: 'mmmm yyyy'};
const timeFormat = {type: 'TIME', pattern: 'hh:mm'};
const palette = {
  dark: {red: .21, green: .16, blue: .29},
  white: {red: 1, green: 1, blue: 1},
  tint: {red: .95, green: .93, blue: .97},
  ink: {red: .18, green: .15, blue: .23}
};

export const workPhotoTypes = [
  ['', 'Эскиз'],
  ['fresh-', 'Свежая татуировка'],
  ['healed-', 'Зажившая татуировка']
];

const dateSerial = value => {
  const ms = Date.parse(value + 'T00:00:00Z');
  return Number.isFinite(ms) ? ms / 86400000 + 25569 : String(value || '');
};
const timeSerial = value => {
  if (!/^\d{2}:\d{2}$/.test(value || '')) return String(value || '');
  const [hours, minutes] = value.split(':').map(Number);
  return (hours * 60 + minutes) / 1440;
};
const amount = value => Number.isFinite(Number(value)) ? Number(value) : 0;

export function reportTables({appointments, clients, expenses}, uploaded) {
  const byId = new Map(clients.map(client => [client.id, client]));
  const totals = new Map();
  function monthFor(date) {
    const key = (date || '').slice(0, 7);
    if (!totals.has(key)) totals.set(key, {received: 0, expected: 0, spent: 0});
    return totals.get(key);
  }
  const sessions = [...appointments].sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time)).map(session => {
    const total = Math.max(0, amount(session.price) - amount(session.discount));
    if (!session.demo) {
      const month = monthFor(session.date);
      if (session.done) month.received += total;
      else if (!session.cancelled) month.expected += total;
    }
    const status = session.demo ? 'Демо' : session.done ? 'Завершён' : session.cancelled ? 'Отменён' : 'Запланирован';
    const links = workPhotoTypes.map(([prefix]) => {
      const fileId = uploaded[prefix + session.id]?.fileId;
      return /^[a-zA-Z0-9_-]{1,128}$/.test(fileId || '') ? {fileId} : '';
    });
    return [dateSerial(session.date), timeSerial(session.time), byId.get(session.clientId)?.name || session.name || '',
      status, amount(session.price), amount(session.discount), total, session.notes || '', ...links];
  });
  const costs = [...expenses].sort((a, b) => a.date.localeCompare(b.date)).map(expense => {
    const cost = amount(expense.amount);
    monthFor(expense.date).spent += cost;
    return [dateSerial(expense.date), expense.note || '', cost];
  });
  const months = [...totals].sort(([a], [b]) => b.localeCompare(a)).map(([month, totals]) =>
    [dateSerial(month + '-01'), totals.received, totals.expected, totals.spent, totals.received - totals.spent]);
  return [
    {title: 'Финансы', headers: ['Месяц', 'Получено', 'Ожидается', 'Расходы', 'Получено − расходы'],
      rows: months, widths: [158, 124, 124, 124, 166], formats: [monthFormat, rubles, rubles, rubles, rubles]},
    {title: 'Сеансы', headers: ['Дата', 'Время', 'Клиент', 'Статус', 'Цена', 'Скидка', 'Итого', 'Описание', ...workPhotoTypes.map(([, label]) => label)],
      rows: sessions, widths: [112, 86, 196, 131, 104, 98, 104, 300, 110, 160, 175],
      formats: [dateFormat, timeFormat, null, null, rubles, rubles, rubles], filter: true},
    {title: 'Расходы', headers: ['Дата', 'Комментарий', 'Сумма'],
      rows: costs, widths: [112, 320, 124], formats: [dateFormat, null, rubles], filter: true}
  ];
}

function makeCell(value, format) {
  if (value?.fileId) {
    return {userEnteredValue: {stringValue: 'Открыть'}, userEnteredFormat: {
      ...format, textFormat: {link: {uri: 'https://drive.google.com/open?id=' + value.fileId}}
    }};
  }
  const userEnteredValue = typeof value === 'number' ? {numberValue: value}
    : {stringValue: String(value ?? '')};
  return {userEnteredValue, userEnteredFormat: format};
}

export function reportRequests(tables, existing) {
  const used = new Set(existing.map(sheet => sheet.sheetId));
  let nextId = 1;
  const blank = existing.length === 1 && !tables.some(table => table.title === existing[0].title) ? existing[0] : null;
  return tables.flatMap((table, index) => {
    const current = existing.find(item => item.title === table.title) || (index === 0 ? blank : null);
    while (used.has(nextId)) nextId++;
    const sheetId = current?.sheetId ?? nextId++;
    const rowCount = Math.max(100, table.rows.length + 1);
    const columnCount = table.headers.length;
    const gridProperties = {rowCount, columnCount, frozenRowCount: 1, hideGridlines: true};
    const properties = {sheetId, title: table.title, gridProperties, tabColorStyle: {rgbColor: palette.dark}};
    const range = {sheetId, startRowIndex: 0, endRowIndex: rowCount, startColumnIndex: 0, endColumnIndex: columnCount};
    const rows = [{values: table.headers.map(label => makeCell(label, {
      backgroundColor: palette.dark, textFormat: {bold: true, foregroundColor: palette.white}, verticalAlignment: 'MIDDLE'
    }))}, ...table.rows.map((row, rowIndex) => ({values: row.map((value, column) => makeCell(value, {
      backgroundColor: rowIndex % 2 ? palette.white : palette.tint,
      textFormat: {foregroundColor: palette.ink}, verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      ...(table.formats[column] ? {numberFormat: table.formats[column]} : {})
    }))}))];
    return [
      ...(current && table.filter ? [{clearBasicFilter: {sheetId}}] : []),
      current ? {updateSheetProperties: {properties, fields: 'title,gridProperties,tabColorStyle'}} : {addSheet: {properties}},
      {updateCells: {range, rows, fields: 'userEnteredValue,userEnteredFormat'}},
      ...table.widths.map((pixelSize, column) => ({updateDimensionProperties: {
        range: {sheetId, dimension: 'COLUMNS', startIndex: column, endIndex: column + 1},
        properties: {pixelSize}, fields: 'pixelSize'
      }})),
      {updateDimensionProperties: {range: {sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1},
        properties: {pixelSize: 34}, fields: 'pixelSize'}},
      ...(table.filter ? [{setBasicFilter: {filter: {range: {...range, endRowIndex: Math.max(2, rows.length)}}}}] : [])
    ];
  });
}
