const fs = require('fs');
const XLSX = require('xlsx');

const DISK_URL = "https://disk.yandex.ru/d/ptTBhsfZMhVHBg";
const FILE_PATH = "/Zebryata_site.xlsx";

function num(v){ if(typeof v==="number")return v; if(v===null||v===undefined||v==="")return 0; return +String(v).replace(/[\s,]/g,"")||0; }
function normName(s){ return String(s||"").trim().toLowerCase().replace(/ё/g,"е").replace(/\s+/g," "); }
function findSheet(wb,variants){
  const wanted = variants.map(normName);
  for(const name of wb.SheetNames) if(wanted.includes(normName(name))) return name;
  return null;
}

function parseXlsx(wb){
  const d = { collected:0, spent:0, fee:0, cats:[], status:{}, holidays:[], months:[] };
  function rows(variants){
    const name = findSheet(wb, variants);
    if(!name) return [];
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, defval:"" });
  }
  rows(["Финансы"]).forEach(r=>{
    if(!r||!r[0]) return;
    const k = String(r[0]).trim().toLowerCase(), v = num(r[1]);
    if(k==="collected"||k==="собрано") d.collected = v;
    if(k==="spent"||k==="потрачено")   d.spent = v;
    if(k==="fee"||k==="сбор")          d.fee = v;
  });
  rows(["Категории"]).forEach(r=>{ if(r&&r[0]&&r[1]!=="") d.cats.push([String(r[0]).trim(), num(r[1])]); });
  rows(["Статусы"]).forEach(r=>{ if(r&&r[0]&&r[1]!=="") d.status[String(r[0]).trim().toLowerCase()] = num(r[1]); });
  rows(["Праздники"]).forEach(r=>{ if(r&&r[0]&&r[1]) d.holidays.push([String(r[0]),String(r[1]),String(r[2]||""),String(r[3]||"")]); });
  rows(["Именинники"]).forEach(r=>{ if(r&&r[0]&&r[1]!=="") d.months.push([String(r[0]).trim(), num(r[1])]); });
  return d;
}

(async () => {
  console.log('Получаем ссылку на скачивание…');
  const ep = `https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=${encodeURIComponent(DISK_URL)}&path=${encodeURIComponent(FILE_PATH)}`;
  const r = await fetch(ep);
  if(!r.ok) throw new Error('Диск ответил ' + r.status);
  const j = await r.json();
  if(!j.href) throw new Error('Нет ссылки на скачивание: ' + JSON.stringify(j));

  console.log('Скачиваем файл…');
  const fileResp = await fetch(j.href);
  if(!fileResp.ok) throw new Error('Скачивание не удалось: ' + fileResp.status);
  const buf = Buffer.from(await fileResp.arrayBuffer());

  console.log('Парсим xlsx…');
  const wb = XLSX.read(buf, { type: 'buffer' });
  console.log('Листы в файле:', wb.SheetNames.join(', '));

  const data = parseXlsx(wb);
  const isEmpty = !(data.collected>0 || data.cats.length || Object.keys(data.status).length || data.holidays.length || data.months.length);
  if(isEmpty){
    console.warn('Данные пустые — data.js не обновляем');
    process.exit(0);
  }
  data.updatedAt = new Date().toISOString();

  const content = '/* Автогенерировано GitHub Actions ' + data.updatedAt + ' */\n' +
                  'window.BUDGET_DATA = ' + JSON.stringify(data, null, 2) + ';\n';
  fs.writeFileSync('data.js', content);
  console.log('data.js обновлён');
})().catch(e => { console.error('Ошибка:', e.message); process.exit(1); });
