/** Google Apps Script, привязанный к существующей таблице Пифагора. */
const SPREADSHEET_ID = '1yDDo05Y91d-EpcqZPfb9YmTVeIan9ons-ZqO89PC4I0';
const SHEET_NAME = 'Кандидаты';
// Вставьте адрес опубликованного сайта после настройки GitHub Pages.
const ALLOWED_ORIGIN = 'https://YOUR_GITHUB_USERNAME.github.io';
function doPost(e) {
  try {
    const p = e.parameter || {};
    const required = ['candidate','subject','interviewer','hours','comment','verdict','knowledge','explanation','communication','reliability'];
    for (const key of required) if (!String(p[key] || '').trim()) throw new Error('Не заполнено поле: ' + key);
    const ratings = ['knowledge','explanation','communication','reliability'].map(k => Number(p[k]));
    if (ratings.some(n => !Number.isInteger(n) || n < 1 || n > 5)) throw new Error('Некорректная оценка');
    const hours = Number(p.hours);
    if (!Number.isFinite(hours) || hours < 0 || hours > 50) throw new Error('Некорректная нагрузка');
    if (!['Рекомендую','Сомневаюсь','Не рекомендую'].includes(p.verdict)) throw new Error('Некорректный вердикт');
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Не найдена вкладка Кандидаты');
    const safe = x => {const s = String(x).trim().slice(0,800); return /^[=+\-@]/.test(s) ? "'" + s : s};
    const lock = LockService.getScriptLock();lock.waitLock(15000);
    try {
      const row = sheet.getLastRow()+1;
      sheet.getRange(row,1,1,12).setValues([[new Date(),safe(p.candidate),safe(p.subject),safe(p.interviewer),...ratings,hours,safe(p.comment),p.verdict,'На рассмотрении']]);
      sheet.getRange(row,13).setFormula(`=AVERAGE(E${row}:H${row})`);
    } finally {lock.releaseLock()}
    return response(true,'Сохранено');
  } catch (err) {return response(false,err.message)}
}
function response(ok,message){
  // Ответ отображается в скрытом iframe и передаёт результат странице GitHub Pages.
  const payload=JSON.stringify({source:'pifagor-interview',ok,message}).replace(/</g,'\\u003c');
  return HtmlService.createHtmlOutput('<script>parent.postMessage('+payload+','+JSON.stringify(ALLOWED_ORIGIN)+');<\\/script>'.replace('<\\/script>','</script>'));
}
