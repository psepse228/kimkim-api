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
const RED = 'C00000';
const BLUE = '1F4E79';
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
  return new TextRun({
    text: String(text),
    font: FONT,
    size: opts.sz || SZ,
    bold: opts.bold || false,
    italics: opts.italic || false,
    color: opts.color || undefined,
  });
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
  'Узбекистан','Узбекистана','Узбекистане','Узбекистану',
  'Россия','России','Российской','Китай','Китая','Китае',
  'Германия','Германии','Франция','Франции','США','Япония','Японии',
  'Корея','Кореи','Индия','Индии','Турция','Турции','ОАЭ','Катар',
  'Казахстан','Казахстана','Азербайджан','Туркменистан','Таджикистан',
  'Кыргызстан','Пакистан','Бангладеш','Сингапур','Лондон','Москва',
  'Москве','Пекин','Пекине','Ташкент','Ташкенте','Париж','Париже',
  'Берлин','Берлине','Дубай','Дубае','Астана','Алматы','Токио',
  'Африка','Африке','Африки','Америка','Америке','Америки',
  'Европа','Европе','Европы','Азия','Азии','Центральной','Южной',
  'Восточной','Западной','Северной','Ближний','Восток','Востоке',
]);

function parseText(text, companyName) {
  if (!text) return [r('')];
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
        // (year) - italic red
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
function tokenizeText(text) {
  if (!text) return [];
  const runs = [];
  // Split by word boundaries but keep spaces and punctuation
  const tokens = text.split(/(\s+)/);
  for (const token of tokens) {
    if (!token) continue;
    const clean = token.replace(/[«».,;:!?()\---]/g, '').trim();
    if (/^\d+$/.test(clean)) {
      runs.push(r(token, { color: RED }));
    } else if (/^[A-Za-z][A-Za-z0-9\-\.]{0,}$/.test(clean) && clean.length > 1) {
      runs.push(r(token, { color: BLUE }));
    } else if (GEO_RU.has(clean)) {
      runs.push(r(token, { color: BLUE }));
    } else {
      runs.push(r(token));
    }
  }
  return runs;
}

// Paragraph helpers
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
  const data = await res.json();
  const text = (data.answer || '') + '\n\n' + (data.results || []).map(r => `${r.title}: ${r.content}`).join('\n\n');
  if (returnSources) return { text, sources: (data.results || []).map(r => ({ title: r.title, url: r.url })) };
  return text;
}

async function researchPerson(input) {
  const { query, title, company, linkedin, website, extraContext } = input;
  const nameQuery = query.includes('http') ? (company || query) : query;
  const coQuery = company || nameQuery;

  const uzbekSearch = await tavilySearch(`${coQuery} Uzbekistan investment projects cooperation agreements 2025`, true);
  const uzbekText = uzbekSearch.text;
  const uzbekSources = uzbekSearch.sources || [];

  const [bioText, _uzbek, newsText, linkedinText, personImg, logoImg] = await Promise.all([
    tavilySearch(`${nameQuery} biography education career professional background 2025`),
    Promise.resolve(uzbekText),
    tavilySearch(`${coQuery} annual report revenue employees assets 2024 2025 latest`),
    linkedin ? tavilySearch(`${nameQuery} LinkedIn career history current role 2025`) : Promise.resolve(''),
    searchImg(`${nameQuery} official portrait professional headshot high resolution`),
    searchImg(`${coQuery} official logo transparent high resolution`),
  ]);

  const context = [
    extraContext ? `КОНТЕКСТ ОТ КЛИЕНТА (наивысший приоритет):\n${extraContext}` : '',
    `БИОГРАФИЯ:\n${bioText}`,
    `ДЕЯТЕЛЬНОСТЬ В УЗБЕКИСТАНЕ:\n${uzbekText}`,
    `НОВОСТИ И ФИНАНСЫ:\n${newsText}`,
    linkedinText ? `LINKEDIN:\n${linkedinText}` : '',
  ].filter(Boolean).join('\n\n---\n\n');

  const prompt = `Ты составляешь официальный справочный документ для Министерства инвестиций Республики Узбекистан. Используй ВСЕ предоставленные данные, особенно контекст от клиента.

${context}
${title ? `\nДолжность: ${title}` : ''}
${company ? `\nКомпания: ${company}` : ''}

Верни ТОЛЬКО JSON. Все поля строго на русском языке. Если данных нет - "Информация не найдена". Никогда null.

{
  "surname_caps": "ФАМИЛИЯ заглавными (только фамилия, латиница если иностранец)",
  "first_name": "Имя с заглавной (только имя)",
  "title": "Полное название должности на русском языке",
  "company_name": "Название компании на языке оригинала (латиница)",
  "company_description_paragraphs": [
    "Параграф 1 (2-3 предложения): основная деятельность, год основания, штаб-квартира (город, страна). Конкретные факты.",
    "Параграф 2 (2-3 предложения): масштаб - число сотрудников, выручка, активы с конкретными цифрами и единицами. Пример: 'Совокупные активы - $ 283,7 млрд. Доход - $ 21,4 млрд. (2024 г.).'",
    "Параграф 3 (2-3 предложения): география деятельности - в скольких странах/регионах работает, ключевые рынки. Все страны и регионы указать.",
    "Параграф 4 (2-3 предложения): ключевые направления деятельности, чем знаменита компания, последние достижения/стратегия."
  ],
  "company_activities_paragraph": "2-3 предложения: перечисли основные направления деятельности компании в виде связного текста, без списков.",
  "company_intro": "1-2 предложения: с какого года и как именно компания сотрудничает с Узбекистаном. Конкретная дата/год.",
  "company_uzbekistan": [
    {
      "sector": "Название отрасли",
      "description": "2-3 конкретных предложения: проекты, суммы инвестиций, партнёры, договорённости, даты. Используй контекст от клиента в первую очередь."
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
- company_description_paragraphs: ровно 4 параграфа, каждый содержательный, с конкретными цифрами
- company_activities_list: 4-6 пунктов, кратко
- company_uzbekistan: минимум 3 отрасли с деталями и цифрами
- education: хронологически (от старого к новому), минимум 1
- career: хронологически от старого к новому, текущая должность - ПОСЛЕДНЯЯ запись, минимум 4
- Годы ТОЛЬКО с дефисом: "1999-2007 гг." НЕ "1999-2007 гг."
- Суммы с единицей: $ 110 млн., € 4,6 млрд.
- Имена компаний всегда в кавычках «»
- Все термины на русском языке`;

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${OPENAI_API_KEY}` },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      max_tokens: 3500,
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
      out.push(para(parseText(paraText, cn), { indent: 'body' }));
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
    out.push(para(parseText(profile.company_intro, cn), { indent: 'body' }));
  }

  for (const s of profile.company_uzbekistan) {
    out.push(para([
      r(s.sector + ': ', { bold: true }),
      ...parseText(s.description, cn),
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
