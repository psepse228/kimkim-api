const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak,
        ImageRun, UnderlineType, NumberFormat } = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type, X-Admin-Key');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const TAVILY_API_KEY = process.env.TAVILY_API_KEY;
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const ADMIN_KEY = process.env.ADMIN_KEY || 'solura-admin-2026';
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ─── CONSTANTS (from TotalEnergies template XML) ──────────────────────────────
const FONT = 'Cambria';
const SZ_TITLE = 56;   // 28pt - main cover heading
const SZ_H2 = 30;      // 15pt - section headings
const SZ = 30;         // 15pt - body
const SZ_TINY = 10;    // 5pt - spacer
const RED = 'C00000';   // numbers, digits, financial figures
const BLUE = '1F4E79';
const YELLOW_HL = 'FFFF00'; // highlight for uncertain/unverified data
const SP = { after: 120, before: 120, line: 288, lineRule: 'auto' };
const SP0 = { after: 0, line: 288, lineRule: 'auto' };
const SP_LIST = { after: 0, line: 288, lineRule: 'auto' };
const SP_TINY = { after: 0, line: 20, lineRule: 'atLeast' };
const INDENT_BODY = { firstLine: 709 };
const INDENT_LIST = { left: 1134 };
const PAGE_W = 11906;
const MARGIN = 1134;
const CW = PAGE_W - MARGIN * 2; // 9638

const BRD = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDS = { top: BRD, bottom: BRD, left: BRD, right: BRD };
const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDS = { top: NO, bottom: NO, left: NO, right: NO };

// ─── TEXT HELPERS ─────────────────────────────────────────────────────────────
function r(text, opts = {}) {
  const run = {
    text: String(text),
    font: FONT,
    size: opts.sz || SZ,
    bold: opts.bold || false,
    italics: opts.italic || false,
    color: opts.color || undefined,
  };
  if (opts.highlight) run.highlight = opts.highlight;
  return new TextRun(run);
}

// Parse year string - digits red, text default
function yearRuns(str) {
  if (!str) return [r('')];
  return String(str).split(/(\d+)/).filter(t => t.length > 0).map(t =>
    /^\d+$/.test(t) ? r(t, { color: RED }) : r(t)
  );
}

// Parse a full text segment with entity coloring
// Rules:
// - «CompanyName» → «black + name BLUE bold + »black
// - (Geo/Year) → italic, geo=blue, year=red
// - Standalone digits → red
// - Latin words/abbreviations → blue
// - Known geo names in Russian → blue
// - Money amounts: digits red bold, unit (млрд/млн/тыс) black bold
const GEO_RU = new Set([
  // Countries
  'Узбекистан','Узбекистана','Узбекистане','Узбекистану','Узбекистаном',
  'Россия','России','Российской','Российской','Российская','Российское','Российское','РФ',
  'Китай','Китая','Китае','Китаем','КНР',
  'Германия','Германии','Германию','Германией',
  'Франция','Франции','Францию','Францией',
  'США','Япония','Японии','Японию','Японией',
  'Корея','Кореи','Корею','Кореей','Южная','Южной',
  'Индия','Индии','Индию','Индией',
  'Турция','Турции','Турцию','Турцией',
  'ОАЭ','Катар','Катара','Катаре','Катаром',
  'Израиль','Израиля','Израиле','Израилем',
  'Казахстан','Казахстана','Казахстане','Казахстаном',
  'Азербайджан','Азербайджана','Азербайджане','Азербайджаном',
  'Туркменистан','Туркменистана','Туркменистане','Туркменистаном',
  'Таджикистан','Таджикистана','Таджикистане','Таджикистаном',
  'Кыргызстан','Кыргызстана','Кыргызстане','Кыргызстаном',
  'Пакистан','Пакистана','Пакистане','Пакистаном',
  'Бангладеш','Бангладеша',
  'Сингапур','Сингапура','Сингапуре','Сингапуром',
  'Нидерланды','Нидерландов','Нидерландах','Нидерландами',
  'Швейцария','Швейцарии','Швейцарию','Швейцарией',
  'Австралия','Австралии','Австралию','Австралией',
  'Канада','Канады','Канаде','Канадой',
  'Бразилия','Бразилии','Бразилию','Бразилией',
  'Великобритания','Великобритании','Великобританию','Великобританией',
  'Ирландия','Ирландии','Ирландию','Ирландией',
  'Норвегия','Норвегии','Норвегию','Норвегией',
  'Швеция','Швеции','Швецию','Швецией',
  'Финляндия','Финляндии','Финляндию','Финляндией',
  'Польша','Польши','Польше','Польшей',
  'Чехия','Чехии','Чехию','Чехией',
  'Австрия','Австрии','Австрию','Австрией',
  'Бельгия','Бельгии','Бельгию','Бельгией',
  'Дания','Дании','Данию','Данией',
  'Испания','Испании','Испанию','Испанией',
  'Португалия','Португалии','Португалию','Португалией',
  'Италия','Италии','Италию','Италией',
  'Румыния','Румынии','Румынию','Румынией',
  'Венгрия','Венгрии','Венгрию','Венгрией',
  'Греция','Греции','Грецию','Грецией',
  'Мексика','Мексики','Мексике','Мексикой',
  'Аргентина','Аргентины','Аргентине','Аргентиной',
  'Колумбия','Колумбии','Колумбию','Колумбией',
  'Чили','Перу','Венесуэла','Венесуэлы',
  'Египет','Египта','Египте','Египтом',
  'ЮАР','Нигерия','Нигерии','Нигерией',
  'Марокко','Алжир','Алжира','Алжире','Алжиром',
  'Ангола','Анголы','Анголе','Анголой',
  'Кения','Кении','Кению','Кенией',
  'Иран','Ирана','Иране','Ираном',
  'Ирак','Ирака','Ираке','Ираком',
  'Саудовская','Саудовской','Саудовскую',
  'Аравия','Аравии','Аравию','Аравией',
  'Кувейт','Кувейта','Кувейте','Кувейтом',
  'Бахрейн','Бахрейна','Бахрейне','Бахрейном',
  'Оман','Омана','Омане','Оманом',
  'Афганистан','Афганистана','Афганистане',
  'Монголия','Монголии','Монголию','Монголией',
  'Мьянма','Таиланд','Таиланда','Таиланде','Таиландом',
  'Вьетнам','Вьетнама','Вьетнаме','Вьетнамом',
  'Индонезия','Индонезии','Индонезию','Индонезией',
  'Малайзия','Малайзии','Малайзию','Малайзией',
  'Филиппины','Филиппин','Филиппинах',
  'Тайвань','Тайваня','Тайване','Тайванем',
  'Армения','Армении','Армению','Арменией',
  'Грузия','Грузии','Грузию','Грузией',
  'Беларусь','Беларуси','Беларусью',
  'Украина','Украины','Украине','Украиной',
  'Молдова','Молдовы','Молдове','Молдовой',
  'Латвия','Латвии','Латвию','Латвией',
  'Литва','Литвы','Литве','Литвой',
  'Эстония','Эстонии','Эстонию','Эстонией',
  'Словакия','Словакии','Словакию','Словакией',
  'Словения','Словении','Словению','Словенией',
  'Хорватия','Хорватии','Хорватию','Хорватией',
  'Сербия','Сербии','Сербию','Сербией',
  'Люксембург','Люксембурга','Люксембурге',
  'Исландия','Исландии','Исландию','Исландией',
  // Cities
  'Лондон','Лондоне','Лондона','Лондоном',
  'Москва','Москве','Москвы','Москвой',
  'Пекин','Пекине','Пекина','Пекином',
  'Ташкент','Ташкенте','Ташкента','Ташкентом',
  'Париж','Париже','Парижа','Парижем',
  'Берлин','Берлине','Берлина','Берлином',
  'Дубай','Дубае','Дубая','Дубаем',
  'Астана','Астане','Астаны','Астаной',
  'Алматы','Токио','Сеул','Сеуле','Сеула',
  'Нью-Йорк','Нью-Йорке','Нью-Йорка','Нью-Йорком',
  'Сан-Франциско','Лос-Анджелес','Лос-Анджелесе','Лос-Анджелеса',
  'Хьюстон','Хьюстоне','Хьюстона','Вашингтон','Вашингтоне','Вашингтона',
  'Чикаго','Сиэтл','Бостон','Бостоне','Бостона',
  'Амстердам','Амстердаме','Амстердама',
  'Брюссель','Брюсселе','Брюсселя',
  'Цюрих','Цюрихе','Цюриха','Женева','Женеве','Женевы',
  'Вена','Вене','Вены','Варшава','Варшаве','Варшавы',
  'Прага','Праге','Праги','Стамбул','Стамбуле','Стамбула',
  'Анкара','Анкаре','Анкары',
  'Тегеран','Тегеране','Тегерана',
  'Эр-Рияд','Эр-Рияде','Эр-Рияда',
  'Абу-Даби','Абу-Даби',
  'Доха','Дохе','Дохи',
  'Мумбаи','Дели','Бангалор','Шанхай','Шанхае','Шанхая',
  'Гонконг','Гонконге','Гонконга',
  'Джакарта','Бангкок','Куала-Лумпур',
  'Сидней','Сиднее','Сиднея',
  'Торонто','Монреаль','Ванкувер',
  'Сан-Паулу','Мехико','Буэнос-Айрес',
  'Найроби','Каир','Лагос',
  'Ереван','Тбилиси','Баку','Бишкек','Душанбе','Ашхабад',
  'Минск','Киев','Рига','Вильнюс','Таллин',
  'Осло','Стокгольм','Хельсинки','Копенгаген',
  'Мадрид','Барселоне','Барселона','Рим','Риме','Рима',
  'Милан','Милане','Милана','Лиссабон',
  'Варшава','Варшаве','Варшавы',
  'Краков','Будапешт','Бухарест',
  'Самарканд','Самарканде','Самарканда',
  'Бухара','Бухаре','Бухары',
  'Андижан','Наманган','Фергана',
  // US States
  'Техас','Техаса','Техасе','Техасом',
  'Калифорния','Калифорнии','Калифорнию','Калифорнией',
  'Флорида','Флориды','Флориде','Флоридой',
  'Нью-Йорк','Невада','Невады','Неваде',
  'Делавэр','Делавэра','Делавэре',
  'Вирджиния','Вирджинии','Вирджинию','Вирджинией',
  'Иллинойс','Иллинойса','Иллинойсе',
  'Вашингтон','Вашингтона','Вашингтоне',
  'Аляска','Аляски','Аляске',
  'Миссури','Огайо','Джорджия','Теннесси','Аризона',
  'Массачусетс','Нью-Джерси','Пенсильвания',
  // Regions
  'Африка','Африке','Африки','Африкой','Субсахарской','Субсахарской',
  'Америка','Америке','Америки','Америкой',
  'Европа','Европе','Европы','Европой',
  'Азия','Азии','Азией',
  'Центральная','Центральной','Центральную',
  'Центральноазиатском','Центральноазиатской','Центральноазиатского',
  'Восточная','Восточной','Восточную',
  'Западная','Западной','Западную',
  'Северная','Северной','Северную',
  'Ближний','Ближнем','Востоке','Востока','Востоком',
  'Тихоокеанском','Тихоокеанского','Тихоокеанской',
  'Латинской','Латинская','Латинскую',
  'Скандинавии','Скандинавия','Скандинавию',
  'Арктике','Арктика','Арктики',
  'Средней','Средняя','Среднего',
  'Персидский','Персидском','Персидского',
  'Каспийский','Каспийском','Каспийского','Каспийского',
]);

function parseText(text, companyName) {
  if (!text) return [r('')];
  // Replace em-dash/en-dash with simple hyphen everywhere
  text = text.replace(/[—–]/g, '-');
  const runs = [];
  const escapedCo = companyName
    ? companyName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    : null;

  // Build combined regex
  const patterns = [
    '«([^»]+)»',                                    // 1: «name»
    '\\(([^)]{1,60})\\)',                            // 2: (parenthetical)
    '(\\$\\s*[\\d][\\d\\s,\\.]*(?:млрд|млн|тыс)?(?:\\.)?)', // 3: $ amount
    '([\\d]+(?:[,\\.][\\d]+)?(?:\\s*(?:млрд|млн|тыс))?(?:\\.)?(?=\\s))', // 4: number+unit
    escapedCo ? `(${escapedCo})` : null,            // 5: plain company name
  ].filter(Boolean).join('|');

  const regex = new RegExp(patterns, 'g');
  let last = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > last) {
      runs.push(...tokenizeText(text.slice(last, match.index)));
    }

    if (match[1] !== undefined) {
      // «CompanyName» - guillemets black, name blue bold
      runs.push(r('«'));
      runs.push(r(match[1], { color: BLUE, bold: true }));
      runs.push(r('»'));
    } else if (match[2] !== undefined) {
      const inner = match[2].trim();
      if (/^\d/.test(inner)) {
        // (year) - italic blue (digits blue)
        runs.push(r('(', { italic: true }));
        runs.push(...inner.split(/(\d+)/).filter(t=>t).map(t =>
          /^\d+$/.test(t) ? r(t, { italic: true, color: RED }) : r(t, { italic: true })
        ));
        runs.push(r(')', { italic: true }));
      } else {
        // (geo) - italic blue
        runs.push(r('(', { italic: true }));
        runs.push(r(inner, { italic: true, color: BLUE }));
        runs.push(r(')', { italic: true }));
      }
    } else if (match[3] !== undefined) {
      // $ amount
      runs.push(...parseMoneyStr(match[3]));
    } else if (match[4] !== undefined) {
      // number+unit
      runs.push(...parseMoneyStr(match[4]));
    } else if (match[5] !== undefined) {
      // plain company name - blue bold
      runs.push(r(match[5], { color: BLUE, bold: true }));
    }

    last = match.index + match[0].length;
  }

  if (last < text.length) runs.push(...tokenizeText(text.slice(last)));
  return runs.length > 0 ? runs : [r(text)];
}

function parseMoneyStr(str) {
  str = str.replace(/[—–]/g, '-');
  // Split: currency symbol | digits/comma | unit word
  const runs = [];
  const unitRe = /\b(млрд|млн|тыс|тысяч|миллиард|миллион|баррелей|тонн|долл|долларов|евро)\b\.?/g;
  const digitRe = /[\d]+(?:[,\.][\d]+)*/g;

  let i = 0;
  const combined = /(\$|€|£|¥|\bдолл\.?|\bевро\b)|(\d[\d,\.]*(?:\s*(?:млрд|млн|тыс))?\.?)|(млрд\.?|млн\.?|тыс\.?)/g;
  let m;
  let last = 0;
  while ((m = combined.exec(str)) !== null) {
    if (m.index > last) runs.push(r(str.slice(last, m.index)));
    if (m[1]) runs.push(r(m[1])); // currency symbol - black
    else if (m[2]) runs.push(r(m[2], { color: RED, bold: true })); // digits - red bold
    else if (m[3]) runs.push(r(m[3], { bold: true })); // unit - black bold
    last = m.index + m[0].length;
  }
  if (last < str.length) runs.push(r(str.slice(last)));
  return runs.length > 0 ? runs : [r(str)];
}

// Tokenize plain text: color geo names, latin words, digits
// Punctuation (.,;:-) always stays black - never colored
function tokenizeText(text) {
  if (!text) return [];
  // Replace all em-dashes and en-dashes with simple hyphen
  text = text.replace(/[—–]/g, '-');
  const runs = [];
  // Split into: words | spaces | punctuation (punctuation always black)
  const tokens = text.split(/(\s+|[.,;:!?()«»"']+)/);
  for (const token of tokens) {
    if (!token) continue;
    // Pure punctuation/spaces - always black
    if (/^[\s.,;:!?()«»"'\-]+$/.test(token)) {
      runs.push(r(token));
      continue;
    }
    const clean = token.trim();
    if (!clean) continue;
    // Digits (pure, with %, or decimal) - red
    if (/^\d[\d,\.]*%?$/.test(clean)) {
      runs.push(r(token, { color: RED }));
    // Latin word (company name, abbreviation, English term) - blue
    } else if (/^[A-Za-z][A-Za-z0-9\-\.]*$/.test(clean) && clean.length > 1) {
      runs.push(r(token, { color: BLUE }));
    // Russian geo name - blue
    } else if (GEO_RU.has(clean)) {
      runs.push(r(token, { color: BLUE }));
    } else {
      runs.push(r(token));
    }
  }
  return runs;
}

// Parse text that may contain [?]...[/?] uncertainty markers from GPT
// Uncertain segments get yellow highlight applied via re-parsing each token
function parseTextUncertain(text, companyName) {
  if (!text) return [r('')];
  // Replace em-dash/en-dash globally
  text = text.replace(/[—–]/g, '-');
  const segments = [];
  const uncertainRe = /\[\?\](.*?)\[\/\?\]/gs;
  let last = 0;
  let m;
  while ((m = uncertainRe.exec(text)) !== null) {
    if (m.index > last) segments.push({ text: text.slice(last, m.index), uncertain: false });
    segments.push({ text: m[1], uncertain: true });
    last = m.index + m[0].length;
  }
  if (last < text.length) segments.push({ text: text.slice(last), uncertain: false });

  if (segments.length === 0 || (segments.length === 1 && !segments[0].uncertain)) {
    return parseText(text, companyName);
  }

  const runs = [];
  for (const seg of segments) {
    if (!seg.uncertain) {
      runs.push(...parseText(seg.text, companyName));
    } else {
      // Tokenize uncertain text directly, adding yellow highlight to each token
      const tokens = seg.text.replace(/[—–]/g, '-').split(/(\s+|[.,;:!?()«»"']+)/);
      for (const token of tokens) {
        if (!token) continue;
        if (/^[\s.,;:!?()«»"'\-]+$/.test(token)) {
          runs.push(new TextRun({ text: token, font: FONT, size: SZ, highlight: 'yellow' }));
        } else {
          const clean = token.trim();
          let color;
          if (/^\d[\d,\.]*%?$/.test(clean)) color = RED;
          else if (/^[A-Za-z][A-Za-z0-9\-\.]*$/.test(clean) && clean.length > 1) color = BLUE;
          else if (GEO_RU.has(clean)) color = BLUE;
          runs.push(new TextRun({ text: token, font: FONT, size: SZ, color, highlight: 'yellow' }));
        }
      }
    }
  }
  return runs.length > 0 ? runs : [r(text)];
}


function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align || AlignmentType.BOTH,
    spacing: opts.spacing || SP,
    indent: opts.indent === 'body' ? INDENT_BODY : opts.indent === 'list' ? INDENT_LIST : undefined,
  });
}

function cPara(children, opts = {}) {
  return para(children, { ...opts, align: AlignmentType.CENTER });
}

function ep(sz) {
  return new Paragraph({
    children: [new TextRun({ text: '', font: FONT, size: sz || SZ })],
    spacing: SP0,
  });
}

function epTiny() {
  return new Paragraph({
    children: [new TextRun({ text: '', font: FONT, size: SZ_TINY })],
    spacing: SP_TINY,
  });
}

function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function tc(paragraphs, width, opts = {}) {
  return new TableCell({
    borders: opts.noBorder ? NO_BORDS : BORDS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: opts.vAlign || VerticalAlign.TOP,
    columnSpan: opts.span || 1,
    children: Array.isArray(paragraphs) ? paragraphs : [paragraphs],
  });
}

// ─── IMAGE FETCHING ───────────────────────────────────────────────────────────
async function fetchImg(url) {
  try {
    const res = await fetch(url, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
      signal: AbortSignal.timeout(10000),
    });
    if (!res.ok) return null;
    const ct = res.headers.get('content-type') || '';
    if (!ct.startsWith('image/')) return null;
    const buf = Buffer.from(await res.arrayBuffer());
    if (buf.length < 1000) return null; // skip tiny/broken images
    const type = ct.includes('png') ? 'png' : ct.includes('gif') ? 'gif' : 'jpg';
    return { buffer: buf, type };
  } catch { return null; }
}

async function searchImg(query) {
  try {
    const res = await fetch('https://api.tavily.com/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        api_key: TAVILY_API_KEY,
        query,
        search_depth: 'basic',
        max_results: 8,
        include_images: true,
      }),
    });
    const data = await res.json();
    for (const url of (data.images || [])) {
      if (!url || url.includes('.svg') || url.includes('.gif')) continue;
      const img = await fetchImg(url);
      if (img && img.buffer.length > 5000) return img; // ensure not a placeholder
    }
  } catch {}
  return null;
}

function makeImgPara(imgData, wPx, hPx) {
  if (!imgData) return null;
  try {
    return cPara([new ImageRun({
      data: imgData.buffer,
      transformation: { width: wPx, height: hPx },
      type: imgData.type,
    })], { spacing: SP0 });
  } catch { return null; }
}

// ─── RESEARCH ─────────────────────────────────────────────────────────────────
async function tavilySearch(query, returnSources = false) {
  try {
    const res = await fetch('https://api.tavily.com/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        api_key: TAVILY_API_KEY,
        query,
        search_depth: 'advanced',
        max_results: 8,
        include_answer: true,
      }),
    });
    if (!res.ok) {
      console.error(`Tavily ${res.status} for: ${query}`);
      if (returnSources) return { text: '', sources: [] };
      return '';
    }
    const data = await res.json();
    const text = (data.answer || '') + '\n\n' + (data.results || []).map(r => `${r.title}: ${r.content}`).join('\n\n');
    if (returnSources) return { text, sources: (data.results || []).map(r => ({ title: r.title, url: r.url })) };
    return text;
  } catch (err) {
    console.error(`Tavily error for "${query}": ${err.message}`);
    if (returnSources) return { text: '', sources: [] };
    return '';
  }
}

async function researchPerson(input) {
  const { query, title, company, linkedin, website, extraContext } = input;
  const nameQuery = query.includes('http') ? (company || query) : query;
  const coQuery = company || nameQuery;

  // Run all company + person searches in parallel for speed and coverage
  const [
    uzbekSearch,
    uzbekSearchRu,
    bioText,
    companyOverview,
    companyFinancials,
    companyEmployees,
    companyNews,
    personImg,
    logoImg,
  ] = await Promise.all([
    tavilySearch(`${coQuery} Uzbekistan investment projects cooperation agreements 2024 2025`, true),
    tavilySearch(`${coQuery} Узбекистан инвестиции проекты соглашения сотрудничество 2024 2025`),
    tavilySearch(`${nameQuery} biography education career professional background history`),
    tavilySearch(`${coQuery} company history founded year headquarters CEO leadership overview`),
    tavilySearch(`${coQuery} annual revenue profit financial results fiscal year 2024 2025`),
    tavilySearch(`${coQuery} number of employees offices countries global operations 2024 2025`),
    tavilySearch(`${coQuery} news latest developments strategy acquisitions deals 2025`),
    searchImg(`${nameQuery} official portrait professional headshot high resolution`),
    searchImg(`${coQuery} official logo transparent high resolution`),
  ]);

  const uzbekText = uzbekSearch.text + '\n\n' + uzbekSearchRu;
  const uzbekSources = uzbekSearch.sources || [];

  // If LinkedIn URL provided, fetch it directly + do name search for completeness
  let linkedinText = '';
  if (linkedin) {
    const [liDirect, liSearch] = await Promise.all([
      tavilySearch(`site:linkedin.com "${nameQuery}" career history positions education`),
      tavilySearch(`${nameQuery} ${company || ''} LinkedIn work experience positions roles 2025`),
    ]);
    linkedinText = liDirect + '\n\n' + liSearch;
  }

  const context = [
    extraContext ? `КОНТЕКСТ ОТ КЛИЕНТА (наивысший приоритет):\n${extraContext}` : '',
    `БИОГРАФИЯ:\n${bioText}`,
    `ОБЗОР КОМПАНИИ (история, основание, штаб-квартира, руководство):\n${companyOverview}`,
    `ФИНАНСОВЫЕ ПОКАЗАТЕЛИ (выручка, прибыль, активы 2024-2025):\n${companyFinancials}`,
    `СОТРУДНИКИ И ГЕОГРАФИЯ (численность персонала, офисы, страны):\n${companyEmployees}`,
    `ПОСЛЕДНИЕ НОВОСТИ 2024-2025:\n${companyNews}`,
    `ДЕЯТЕЛЬНОСТЬ В УЗБЕКИСТАНЕ:\n${uzbekText}`,
    linkedinText ? `LINKEDIN:\n${linkedinText}` : '',
  ].filter(Boolean).join('\n\n---\n\n');

  const prompt = `Ты составляешь официальный справочный документ для Министерства инвестиций Республики Узбекистан. Используй ВСЕ предоставленные данные, особенно контекст от клиента. Все разделы о компании должны содержать САМУЮ СВЕЖУЮ информацию из поисковых результатов 2024-2025 года - численность сотрудников, финансовые показатели, офисы, руководство.

${context}
${title ? `\nДолжность: ${title}` : ''}
${company ? `\nКомпания: ${company}` : ''}

Верни ТОЛЬКО JSON. Все поля строго на русском языке. Если данных нет - "Информация не найдена". Никогда null.

{
  "surname_caps": "ФАМИЛИЯ заглавными буквами на РУССКОМ языке (транслитерация если иностранец: John Smith → СМИТ, François Dupont → ДЮПОН). Только фамилия.",
  "first_name": "Имя с заглавной буквы на РУССКОМ языке (транслитерация если иностранец: John → Джон, François → Франсуа). Только имя.",
  "title": "Полное название должности на русском языке",
  "company_name": "Название компании на языке оригинала (латиница)",
  "company_description_paragraphs": [
    "Параграф 1 (3-4 предложения): полная история компании - кто основал, когда, где, с чего начиналась, как развивалась до сегодняшнего дня. Штаб-квартира с конкретным городом и штатом/страной. Используй данные из раздела ОБЗОР КОМПАНИИ.",
    "Параграф 2 (3-4 предложения): АКТУАЛЬНЫЕ финансовые показатели из официальных отчётов 2024-2025 гг. Используй данные из раздела ФИНАНСОВЫЕ ПОКАЗАТЕЛИ. ТОЛЬКО реальные цифры подтверждённые источниками. Оберни КАЖДУЮ неподтверждённую цифру в теги: [?]сумма[/?]. Пример подтверждённого: 'Годовая выручка за 2024 г. составила $ 21,4 млрд.' Пример неподтверждённого: '[?]Выручка оценивается в $ 23 млрд.[/?]'",
    "Параграф 3 (3-4 предложения): численность сотрудников (актуальная цифра из раздела СОТРУДНИКИ И ГЕОГРАФИЯ), география присутствия - конкретно в каких странах, штатах, регионах работает, сколько офисов/объектов.",
    "Параграф 4 (3-4 предложения): ключевые продукты, технологии, услуги - чем конкретно занимается и чем известна в своей отрасли.",
    "Параграф 5 (3-4 предложения): последние новости и достижения 2024-2025 из раздела ПОСЛЕДНИЕ НОВОСТИ, стратегия развития, крупные сделки или проекты."
  ],
  "company_activities_paragraph": "2-3 предложения: перечисли основные направления деятельности компании в виде связного текста, без списков.",
  "company_intro": "1-2 предложения: с какого года и как именно компания сотрудничает с Узбекистаном. Только реальные факты. Если нет подтверждённых данных - напиши об этом честно.",
  "company_uzbekistan": [
    {
      "sector": "Название отрасли",
      "description": "2-3 конкретных предложения. ТОЛЬКО реальные подтверждённые факты: проекты, суммы инвестиций, партнёры, договорённости, даты. Используй контекст от клиента в первую очередь. Неподтверждённые факты оборачивай в [?]...[/?]. Если реальных данных нет совсем - честно укажи отсутствие присутствия и опиши потенциальные направления сотрудничества."
    }
  ],
  "education": [
    {
      "years": "ГГГГ-ГГГГ гг.",
      "institution": "Полное название учебного заведения",
      "country": "Страна",
      "degree": "Степень и специальность на русском"
    }
  ],
  "career": [
    {
      "years": "ГГГГ-ГГГГ гг. или с ГГГГ г.",
      "company": "Название организации",
      "country": "Страна",
      "role": "Должность на русском"
    }
  ]
}

ТРЕБОВАНИЯ:
- ТОНАЛЬНОСТЬ: строго нейтральная, без превосходных степеней ("ведущий", "крупнейший", "выдающийся"), без оценочных суждений - только факты
- ТОЧНОСТЬ ДАННЫХ: используй ТОЛЬКО цифры из предоставленных источников. Если цифра не подтверждена - оберни её в [?]...[/?]. Не придумывай финансовые показатели
- company_description_paragraphs: РОВНО 5 параграфов, КАЖДЫЙ содержательный, ПОЛНЫЕ предложения:
  * Параграф 1: основная деятельность, год основания, штаб-квартира (город, штат/страна). Полная история
  * Параграф 2: финансовые показатели - только реальные цифры. Неподтверждённые в [?]...[/?]
  * Параграф 3: география - конкретно в каких странах, регионах, штатах работает, сколько офисов
  * Параграф 4: ключевые продукты/услуги/технологии - чем известна, что производит
  * Параграф 5: стратегия, последние достижения 2024-2025, планы развития
- company_activities_paragraph: 3-4 предложения в виде связного текста без списков
- company_uzbekistan: минимум 3 раздела. Только реальные данные. Если присутствия нет - чётко об этом написать, затем предложить потенциальные направления
- education: LinkedIn - ГЛАВНЫЙ источник. Хронологически от старого к новому. Включать ВСЕ ступени: бакалавр, магистр, PhD, курсы, сертификаты - всё без исключения
- career: LinkedIn - ГЛАВНЫЙ И ОБЯЗАТЕЛЬНЫЙ источник. Включать КАЖДУЮ позицию из LinkedIn: даже краткосрочные, фриланс, консультации, part-time. Хронологически от старого к новому, текущая - ПОСЛЕДНЯЯ. Не пропускать ни одну строку из LinkedIn
- surname_caps и first_name: ВСЕГДА транслитерировать на русский язык. Примеры: John Smith → СМИТ Джон, Peter Saks → САКС Питер, François → Франсуа, Michael → Майкл
- Годы ТОЛЬКО с дефисом: "1999-2007 гг." НИКАКИХ длинных тире (—) нигде в тексте
- Суммы: $ 110 млн., € 4,6 млрд. - запятая и точка чёрные, не окрашивать
- Имена компаний всегда в кавычках «»
- Города и штаты США писать по-русски: Техас, Калифорния, Нью-Йорк, Хьюстон и т.д.
- Все термины на русском языке
- НИКАКИХ буллет-поинтов или списков с "-" в тексте - только параграфы
- НИКАКИХ длинных тире (—) — везде только дефис (-)`;

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${OPENAI_API_KEY}` },
    body: JSON.stringify({
      model: 'gpt-4o',
      max_tokens: 8000,
      response_format: { type: 'json_object' },
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  const data = await res.json();
  if (!data.choices?.[0]?.message?.content) throw new Error('OpenAI returned no content');
  const profile = JSON.parse(data.choices[0].message.content);
  return { ...profile, personImg, logoImg, uzbekSources };
}

function validate(p) {
  for (const f of ['surname_caps','first_name','title','company_name',
                    'company_description_paragraphs','company_uzbekistan','education','career']) {
    if (!p[f]) throw new Error(`Missing: ${f}`);
  }
  return p;
}

// ─── COVER PAGE ───────────────────────────────────────────────────────────────
function buildCoverPage(profiles) {
  const out = [];

  // "ИНФОРМАЦИЯ" - RED bold sz=56, then line break
  // "к встрече с представителями" - black sz=56
  // "«CompanyName»" - blue bold sz=56
  const companyName = profiles[0]?.company_name || '';

  out.push(cPara([
    r('ИНФОРМАЦИЯ', { bold: true, color: RED, sz: SZ_TITLE }),
    new TextRun({ break: 1, font: FONT, size: SZ_TITLE }),
    r('к встрече с представителями', { sz: SZ_TITLE }),
    new TextRun({ break: 1, font: FONT, size: SZ_TITLE }),
    r('компании ', { sz: SZ_TITLE }),
    new TextRun({ break: 1, font: FONT, size: SZ_TITLE }),
    r('«', { bold: true, sz: SZ_TITLE }),
    r(companyName, { bold: true, color: BLUE, sz: SZ_TITLE }),
    r('»', { bold: true, sz: SZ_TITLE }),
  ], { spacing: SP0 }));

  return out;
}

// ─── COMPANY INFO SECTION ─────────────────────────────────────────────────────
function buildCompanySection(profile) {
  const out = [];
  const cn = profile.company_name;

  // Section heading
  out.push(cPara([
    r('Информация о деятельности компании', { bold: true }),
    new TextRun({ break: 1, font: FONT, size: SZ }),
    r('«', { bold: true }),
    r(cn, { bold: true, color: BLUE }),
    r('»', { bold: true }),
  ], { spacing: SP0 }));

  // Logo centered - large (150×60px approx)
  if (profile.logoImg) {
    out.push(ep());
    const lp = makeImgPara(profile.logoImg, 180, 72);
    if (lp) out.push(lp);
  }
  out.push(ep());

  // Description paragraphs - full page, justified, first-line indent
  const paras = Array.isArray(profile.company_description_paragraphs)
    ? profile.company_description_paragraphs
    : [profile.company_description_paragraphs || ''];

  for (const paraText of paras) {
    if (paraText && paraText.trim()) {
      out.push(para(parseTextUncertain(paraText, cn), { indent: 'body' }));
    }
  }

  // Activities bullet list
  if (profile.company_activities_list?.length) {
    out.push(ep());
    out.push(para([r('Основные направления деятельности:', { bold: true })], { indent: 'body' }));
    for (const item of profile.company_activities_list) {
      out.push(new Paragraph({
        children: parseText(item, cn),
        alignment: AlignmentType.BOTH,
        spacing: SP_LIST,
        indent: INDENT_LIST,
        indent: INDENT_LIST,
      }));
    }
  }

  return out;
}

// ─── UZBEKISTAN SECTION ───────────────────────────────────────────────────────
function buildUzbekSection(profile) {
  const out = [];
  const cn = profile.company_name;

  out.push(cPara([
    r('Информация о деятельности компании', { bold: true }),
    new TextRun({ break: 1, font: FONT, size: SZ }),
    r('«', { bold: true }),
    r(cn, { bold: true, color: BLUE }),
    r('» в Республике Узбекистан', { bold: true }),
  ], { spacing: SP0 }));

  out.push(ep());

  if (profile.company_intro) {
    out.push(para(parseTextUncertain(profile.company_intro, cn), { indent: 'body' }));
  }

  for (const s of profile.company_uzbekistan) {
    out.push(para([
      r(s.sector + ': ', { bold: true }),
      ...parseTextUncertain(s.description, cn),
    ], { indent: 'body' }));
  }

  return out;
}

// ─── PARTICIPANTS TABLE ───────────────────────────────────────────────────────
function formatDate(d, t) {
  const months = ['января','февраля','марта','апреля','мая','июня',
                  'июля','августа','сентября','октября','ноября','декабря'];
  const dt = new Date(d);
  return `${dt.getDate()} ${months[dt.getMonth()]} т.г., ${t}`;
}

function uzbekCell(raw, width) {
  if (!raw) return tc([ep()], width);
  const [nameLine, ...rest] = raw.split('\n');
  const words = nameLine.trim().split(/\s+/);
  const surname = words[0] || '';
  const firstname = words.slice(1).join(' ');
  const title = rest.join(' ');
  return tc([
    new Paragraph({
      children: [
        r(surname.toUpperCase(), { bold: true, color: BLUE }),
        ...(firstname ? [new TextRun({ break: 1, font: FONT, size: SZ }), r(firstname, { color: BLUE })] : []),
      ],
      spacing: SP0,
    }),
    epTiny(),
    ...(title ? [new Paragraph({ children: [r(title)], spacing: SP0, alignment: AlignmentType.BOTH })] : []),
  ], width);
}

function foreignCell(profile, width) {
  if (!profile) return tc([ep()], width);
  return tc([
    new Paragraph({
      children: [
        r(profile.surname_caps, { bold: true, color: BLUE }),
        new TextRun({ break: 1, font: FONT, size: SZ }),
        r(profile.first_name, { color: BLUE }),
      ],
      spacing: SP0,
    }),
    epTiny(),
    new Paragraph({ children: [r(profile.title)], spacing: SP0, alignment: AlignmentType.BOTH }),
  ], width);
}

function buildParticipantsTable(date, time, uzbeks, profiles) {
  const out = [];
  out.push(cPara([r('СПИСОК УЧАСТНИКОВ', { bold: true })], { spacing: SP0 }));
  out.push(ep());

  const dateStr = formatDate(date, time);
  const numW = 400;
  const colW = Math.floor((CW - numW) / 2);

  const rows = [
    new TableRow({ children: [
      tc([para([r('Дата: ', { bold: true }), ...yearRuns(dateStr)])], CW, { span: 3 }),
    ]}),
    new TableRow({ children: [tc([ep()], numW), tc([ep()], colW), tc([ep()], colW)] }),
    new TableRow({ children: [
      tc([ep()], numW),
      tc([cPara([r('От узбекской стороны', { bold: true })], { spacing: SP0 })], colW),
      tc([cPara([r('От иностранной стороны', { bold: true })], { spacing: SP0 })], colW),
    ]}),
    ...[0,1,2,3].map(i => new TableRow({ children: [
      tc([cPara([r(`${i+1}.`, { bold: true })], { spacing: SP0 })], numW),
      uzbekCell(uzbeks[i] || '', colW),
      foreignCell(profiles[i] || null, colW),
    ]})),
    new TableRow({ children: [tc([ep()], numW), tc([ep()], colW), tc([ep()], colW)] }),
  ];

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [numW, colW, colW], rows }));
  return out;
}

// ─── BIO SECTION ──────────────────────────────────────────────────────────────
function buildBioSection(profile) {
  const out = [];
  out.push(pb());

  // Name: SURNAME (blue bold) + space + FIRSTNAME (blue bold) - on one line
  // Then line break, then title (black bold), then «Company» (blue bold)
  out.push(cPara([
    r(profile.surname_caps + ' ', { bold: true, color: BLUE }),
    r(profile.first_name.toUpperCase(), { bold: true, color: BLUE }),
    new TextRun({ break: 1, font: FONT, size: SZ }),
    r(profile.title, { bold: true }),
    new TextRun({ break: 1, font: FONT, size: SZ }),
    r('компании', { bold: true }),
    new TextRun({ break: 1, font: FONT, size: SZ }),
    r('«', { bold: true }),
    r(profile.company_name, { bold: true, color: BLUE }),
    r('»', { bold: true }),
  ], { spacing: SP0 }));

  // Photo centered - full size, not compressed (200×250px)
  if (profile.personImg) {
    out.push(ep());
    const pp = makeImgPara(profile.personImg, 200, 250);
    if (pp) out.push(pp);
  }
  out.push(ep());

  // Bio tables - NO BORDERS (transparent)
  const col1 = 2269;
  const col2 = 280;
  const col3 = CW - col1 - col2;

  // Education table
  const eduRows = [
    new TableRow({ children: [
      tc([para([r('Образование:', { bold: true })], { spacing: SP })], CW, { span: 3, noBorder: true }),
    ]}),
    ...profile.education.map(e => {
      const desc = [
        ...(e.degree ? [r(e.degree)] : []),
        ...(e.institution ? [r(', '), r('«'), r(e.institution, { color: BLUE }), r('»')] : []),
        ...(e.country ? [r(' (', { italic: true }), r(e.country, { italic: true, color: BLUE }), r(')', { italic: true })] : []),
      ];
      return new TableRow({ children: [
        tc([para(yearRuns(e.years), { spacing: SP })], col1, { noBorder: true }),
        tc([para([r('-')], { spacing: SP, align: AlignmentType.CENTER })], col2, { noBorder: true }),
        tc([para(desc.length ? desc : [r(e.degree || '')], { spacing: SP })], col3, { noBorder: true }),
      ]});
    }),
  ];
  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: eduRows }));
  out.push(ep());

  // Career table - chronological, oldest first, current role LAST
  const sortedCareer = [...profile.career].sort((a, b) => {
    const getYear = s => parseInt((s || '').match(/\d{4}/)?.[0] || '0');
    return getYear(a.years) - getYear(b.years);
  });

  const careerRows = [
    new TableRow({ children: [
      tc([para([r('Профессиональная деятельность:', { bold: true })], { spacing: SP })], CW, { span: 3, noBorder: true }),
    ]}),
    ...sortedCareer.map(c => {
      const desc = [
        ...(c.role ? [r(c.role)] : []),
        ...(c.company ? [r(' в компании «'), r(c.company, { color: BLUE }), r('»')] : []),
        ...(c.country ? [r(' (', { italic: true }), r(c.country, { italic: true, color: BLUE }), r(')', { italic: true })] : []),
      ];
      return new TableRow({ children: [
        tc([para(yearRuns(c.years), { spacing: SP })], col1, { noBorder: true }),
        tc([para([r('-')], { spacing: SP, align: AlignmentType.CENTER })], col2, { noBorder: true }),
        tc([para(desc.length ? desc : [r(c.role || '')], { spacing: SP })], col3, { noBorder: true }),
      ]});
    }),
  ];
  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: careerRows }));

  return out;
}

// ─── DOCX ASSEMBLY ────────────────────────────────────────────────────────────
async function generateDocx(profiles, date, time, uzbeks) {
  let children = [];

  // Cover page
  children = children.concat(buildCoverPage(profiles));
  children.push(pb());

  // Company + Uzbekistan sections per profile
  for (let i = 0; i < profiles.length; i++) {
    children = children.concat(buildCompanySection(profiles[i]));
    children.push(pb());
    children = children.concat(buildUzbekSection(profiles[i]));
    if (i < profiles.length - 1) children.push(pb());
  }

  children.push(pb());

  // Participants table
  children = children.concat(buildParticipantsTable(date, time, uzbeks, profiles));

  // Bio sections
  for (const p of profiles) {
    children = children.concat(buildBioSection(p));
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: 16838 },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      children,
    }],
  });
  return await Packer.toBuffer(doc);
}

// ─── USER ─────────────────────────────────────────────────────────────────────
async function upsertUser(user_id, username, full_name) {
  await supabase.from('users').upsert(
    { user_id: String(user_id), username: username || null, full_name: full_name || null, updated_at: new Date().toISOString() },
    { onConflict: 'user_id' }
  );
}

// ─── ENDPOINTS ────────────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  const { inputs, meeting_date, meeting_time, uzbek_participants, user_id, username, full_name, extra_context } = req.body;

  if (!inputs?.length) return res.status(400).json({ error: 'inputs required' });
  if (!meeting_date) return res.status(400).json({ error: 'meeting_date required' });
  if (!uzbek_participants?.length) return res.status(400).json({ error: 'uzbek_participants required' });

  try {
    await upsertUser(user_id, username, full_name);

    const enriched = inputs.map(i => ({ ...i, extraContext: extra_context || null }));

    console.log(`Researching ${inputs.length} person(s)...`);
    const rawProfiles = await Promise.all(enriched.map(i => researchPerson(i)));
    const profiles = rawProfiles.map(validate);
    console.log('Validated. Generating DOCX...');

    const buf = await generateDocx(profiles, meeting_date, meeting_time || '11:00', uzbek_participants);

    const doc_id = `doc_${Date.now()}_${user_id}`;
    const fileName = `${doc_id}.docx`;

    const { error: upErr } = await supabase.storage.from('documents').upload(fileName, buf, {
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      upsert: true,
    });
    if (upErr) throw new Error(`Upload failed: ${upErr.message}`);

    // Save Uzbekistan sources as a separate JSON file
    let allSources = profiles.flatMap((p, i) => (p.uzbekSources || []).map(s => ({
      person: inputs[i]?.query || '',
      company: p.company_name || '',
      title: s.title,
      url: s.url,
    })));
    if (allSources.length) {
      const sourcesFileName = doc_id + '_sources.json';
      await supabase.storage.from('documents').upload(
        sourcesFileName,
        Buffer.from(JSON.stringify(allSources, null, 2)),
        { contentType: 'application/json', upsert: true }
      );
    }

    const download_url = `${SUPABASE_URL}/storage/v1/object/public/documents/${fileName}`;

    await supabase.from('documents').insert({
      doc_id, user_id: String(user_id || 'unknown'), download_url, meeting_date,
      profiles_count: profiles.length,
      queries: inputs.map(i => i.query),
      created_at: new Date().toISOString(),
    });

    console.log('Done:', doc_id);
    if (!allSources) allSources = [];
    const sources_url = allSources.length
      ? `${SUPABASE_URL}/storage/v1/object/public/documents/${doc_id}_sources.json`
      : null;
    return res.json({ success: true, doc_id, download_url, sources_url, meeting_date });
  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/profile/:user_id', async (req, res) => {
  try {
    const { data: user } = await supabase.from('users').select('*').eq('user_id', req.params.user_id).single();
    const { data: docs } = await supabase.from('documents').select('*').eq('user_id', req.params.user_id).order('created_at', { ascending: false }).limit(20);
    return res.json({ user, documents: docs || [] });
  } catch (err) { return res.status(500).json({ error: err.message }); }
});

app.get('/admin/documents', async (req, res) => {
  if (req.headers['x-admin-key'] !== ADMIN_KEY) return res.status(401).json({ error: 'Unauthorized' });
  const { data } = await supabase.from('documents').select('*, users(full_name,username)').order('created_at', { ascending: false }).limit(100);
  return res.json({ documents: data });
});

app.get('/admin/users', async (req, res) => {
  if (req.headers['x-admin-key'] !== ADMIN_KEY) return res.status(401).json({ error: 'Unauthorized' });
  const { data } = await supabase.from('users').select('*').order('created_at', { ascending: false });
  return res.json({ users: data });
});

app.get('/health', (_, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`KimKim API running on port ${PORT}`));
