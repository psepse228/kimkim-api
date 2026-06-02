const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak,
        ImageRun } = require('docx');
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

// ─── STYLE CONSTANTS ──────────────────────────────────────────────────────────
const FONT = 'Cambria';
const SZ = 30;         // 15pt
const SZ_TINY = 10;    // 5pt spacer
const RED = 'C00000';
const BLUE = '1F4E79';
const BLACK = '000000';
const SP = { after: 120, before: 120, line: 288, lineRule: 'auto' };
const SP0 = { after: 0, line: 240, lineRule: 'auto' };
const SP_TINY = { after: 0, line: 20, lineRule: 'atLeast' };
const INDENT = { firstLine: 720 };
const PAGE_W = 11906;
const MARGIN = 1134;
const CW = PAGE_W - MARGIN * 2;

const BRD = { style: BorderStyle.SINGLE, size: 4, color: BLACK };
const BORDS = { top: BRD, bottom: BRD, left: BRD, right: BRD };
const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDS = { top: NO, bottom: NO, left: NO, right: NO };

// Known blue-colored entity patterns: geo names, abbreviations, English words
const GEO_PATTERN = /\b(США|КНР|ЕС|МВФ|ВТО|ООН|АБР|ЕБРР|МФК|ВБ|ОПЕК|СНГ|ЕАЭС|ШОС|Узбекистан[а-я]*|Россия|России|Китай|Китая|Германия|Германии|Франция|Франции|США|Японии|Япония|Корея|Кореи|Индия|Индии|Турция|Турции|ОАЭ|Катар|Лондон[а-я]*|Москв[а-я]*|Пекин[а-я]*|Ташкент[а-я]*|Нью-Йорк[а-я]*|Токио|Париж[а-я]*|Берлин[а-я]*|Дубай|Астана|Алмат[а-я]*|Центральн[а-я]+ Азии?|Центральная Азия)\b/g;
const ABBR_PATTERN = /\b([A-Z]{2,})\b/g;
const LATIN_PATTERN = /\b([A-Za-z][A-Za-z0-9\-\.]{2,})\b/g;

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

// Parse year string — digits red, rest black
function yearRuns(str) {
  return str.split(/(\d+)/).filter(t => t.length > 0).map(t =>
    /^\d+$/.test(t) ? r(t, { color: RED, bold: false }) : r(t)
  );
}

// Parse monetary amount — digits/numbers red bold, unit words black bold, currency black
function moneyRuns(str) {
  // e.g. "$ 4,6 млрд." or "110 млн долларов"
  return str.split(/(\d[\d\s,\.]*(?:млрд|млн|тыс|тысяч|миллиард|миллион)?\.?)/).map((t, i) => {
    if (i % 2 === 1) return r(t, { color: RED, bold: true });
    // unit words bold
    if (/^(млрд|млн|тыс|тысяч|миллиард|миллион)/.test(t.trim())) return r(t, { bold: true });
    return r(t);
  });
}

// Main text parser — handles colors for all entity types
function parseText(text, companyName) {
  if (!text) return [r('')];
  const runs = [];

  // Build regex: «...» | money | (geo/year) | digits | company plain | geo words | latin words
  const escapedCo = companyName ? companyName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') : null;
  const parts = [
    '«([^»]+)»',                                           // group 1: «company»
    '(\\$\\s*[\\d][\\d\\s,\\.]*(?:млрд|млн|тыс)?[\\.]?)', // group 2: $ amounts
    '\\(([^)]+)\\)',                                        // group 3: (parenthetical)
    '(\\d+[\\d\\s,\\.]*(?:млрд|млн|тыс)?[\\.]?)',          // group 4: numbers/amounts
    escapedCo ? `(${escapedCo})` : null,                   // group 5: company name plain
  ].filter(Boolean).join('|');

  const regex = new RegExp(parts, 'g');
  let last = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    // Plain text before match — color geo/latin entities
    if (match.index > last) {
      const plain = text.slice(last, match.index);
      runs.push(...colorPlainText(plain));
    }

    if (match[1] !== undefined) {
      // «company name» — blue bold
      runs.push(r('«'));
      runs.push(r(match[1], { color: BLUE, bold: true }));
      runs.push(r('»'));
    } else if (match[2] !== undefined) {
      // $ amount — $ black, digits red bold, unit bold
      runs.push(...moneyRuns(match[2]));
    } else if (match[3] !== undefined) {
      // (parenthetical)
      const inner = match[3];
      if (/^\d/.test(inner)) {
        runs.push(r('(', { italic: true }));
        runs.push(r(inner, { italic: true, color: RED }));
        runs.push(r(')', { italic: true }));
      } else {
        runs.push(r('(', { italic: true }));
        runs.push(r(inner, { italic: true, color: BLUE }));
        runs.push(r(')', { italic: true }));
      }
    } else if (match[4] !== undefined) {
      // standalone numbers — red
      const numStr = match[4];
      if (/млрд|млн|тыс/.test(numStr)) {
        runs.push(...moneyRuns(numStr));
      } else {
        runs.push(...numStr.split(/(\d+)/).filter(t => t).map(t =>
          /^\d+$/.test(t) ? r(t, { color: RED }) : r(t)
        ));
      }
    } else if (match[5] !== undefined) {
      // plain company name mention — blue bold
      runs.push(r(match[5], { color: BLUE, bold: true }));
    }

    last = match.index + match[0].length;
  }

  if (last < text.length) runs.push(...colorPlainText(text.slice(last)));
  return runs.length > 0 ? runs : [r(text)];
}

// Color plain text: geo names blue, latin words blue, digits red
function colorPlainText(text) {
  if (!text) return [];
  const result = [];
  // Simple tokenize by spaces and punctuation
  const tokens = text.split(/(\s+|(?=[«»(),;:!?])|(?<=[«»(),;:!?]))/);
  for (const token of tokens) {
    if (!token) continue;
    if (/^\d+$/.test(token)) {
      result.push(r(token, { color: RED }));
    } else if (/^[A-Za-z][A-Za-z0-9\-\.]{1,}$/.test(token)) {
      // Latin word — blue
      result.push(r(token, { color: BLUE }));
    } else if (GEO_PATTERN.test(token)) {
      GEO_PATTERN.lastIndex = 0;
      result.push(r(token, { color: BLUE }));
    } else {
      result.push(r(token));
    }
  }
  return result;
}

function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align || AlignmentType.BOTH,
    spacing: opts.spacing || SP,
    indent: opts.indent ? INDENT : undefined,
  });
}

function cPara(children, opts = {}) {
  return para(children, { ...opts, align: AlignmentType.CENTER });
}

function ep(sz) {
  return new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: sz || SZ })], spacing: SP0 });
}
function epTiny() {
  return new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: SZ_TINY })], spacing: SP_TINY });
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
      headers: { 'User-Agent': 'Mozilla/5.0' },
      signal: AbortSignal.timeout(8000),
    });
    if (!res.ok) return null;
    const ct = res.headers.get('content-type') || '';
    if (!ct.includes('image')) return null;
    const buf = Buffer.from(await res.arrayBuffer());
    const type = ct.includes('png') ? 'png' : ct.includes('gif') ? 'gif' : 'jpg';
    return { buffer: buf, type };
  } catch { return null; }
}

async function searchImg(query) {
  try {
    const res = await fetch('https://api.tavily.com/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ api_key: TAVILY_API_KEY, query, search_depth: 'basic', max_results: 5, include_images: true }),
    });
    const data = await res.json();
    for (const url of (data.images || [])) {
      if (!url || url.includes('.svg')) continue;
      const img = await fetchImg(url);
      if (img) return img;
    }
  } catch {}
  return null;
}

function imgPara(imgData, wCm, hCm) {
  if (!imgData) return null;
  try {
    return cPara([new ImageRun({
      data: imgData.buffer,
      transformation: { width: Math.round(wCm * 28.35), height: Math.round(hCm * 28.35) },
      type: imgData.type,
    })], { spacing: SP0 });
  } catch { return null; }
}

// ─── RESEARCH ─────────────────────────────────────────────────────────────────
async function tavilySearch(query) {
  const res = await fetch('https://api.tavily.com/search', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      api_key: TAVILY_API_KEY, query,
      search_depth: 'advanced', max_results: 8,
      include_answer: true,
    }),
  });
  const data = await res.json();
  return (data.answer || '') + '\n\n' + (data.results || []).map(r => `${r.title}: ${r.content}`).join('\n\n');
}

async function researchPerson(input) {
  const { query, title, company, linkedin, website, file, fileName, extraContext } = input;

  // Build search queries
  const nameQuery = query.includes('http') ? company || query : query;
  const companyQuery = company || nameQuery;

  const searches = [
    tavilySearch(`${nameQuery} biography education career background`),
    tavilySearch(`${companyQuery} Uzbekistan investment projects cooperation 2024 2025`),
    tavilySearch(`${companyQuery} latest news 2025`),
    linkedin ? tavilySearch(`site:linkedin.com "${nameQuery}"`) : Promise.resolve(''),
    website ? tavilySearch(`site:${website.replace(/https?:\/\//, '')} about leadership`).catch(() => '') : Promise.resolve(''),
    searchImg(`${nameQuery} portrait photo professional`),
    searchImg(`${companyQuery} company logo official`),
  ];

  const [personText, uzbekText, newsText, linkedinText, websiteText, personImg, logoImg] = await Promise.all(searches);

  // Combine all context
  const allContext = [
    extraContext ? `ДОПОЛНИТЕЛЬНЫЙ КОНТЕКСТ ОТ КЛИЕНТА:\n${extraContext}` : '',
    file ? `ЗАГРУЖЕННЫЙ ДОКУМЕНТ (${fileName}):\n[Содержимое предоставлено пользователем]` : '',
    `БИОГРАФИЯ И КАРЬЕРА:\n${personText}`,
    `ДЕЯТЕЛЬНОСТЬ В УЗБЕКИСТАНЕ:\n${uzbekText}`,
    `ПОСЛЕДНИЕ НОВОСТИ:\n${newsText}`,
    linkedinText ? `LINKEDIN:\n${linkedinText}` : '',
    websiteText ? `САЙТ КОМПАНИИ:\n${websiteText}` : '',
  ].filter(Boolean).join('\n\n---\n\n');

  const prompt = `Ты — эксперт по составлению официальных справочных материалов для Министерства инвестиций Республики Узбекистан.

Все данные ниже — используй ВСЁ что найдёшь. Приоритет: сначала "ДОПОЛНИТЕЛЬНЫЙ КОНТЕКСТ ОТ КЛИЕНТА" и "ЗАГРУЖЕННЫЙ ДОКУМЕНТ", потом остальные источники.

${allContext}

${title ? `Известная должность: ${title}` : ''}
${company ? `Известная компания: ${company}` : ''}

Верни ТОЛЬКО JSON. Все поля на русском. Если информация не найдена — "Информация не найдена". Никогда null или пустые строки.

{
  "surname_caps": "ФАМИЛИЯ заглавными (только фамилия)",
  "first_name": "Имя с заглавной буквы (только имя)",
  "full_name_bio": "ИМЯ ФАМИЛИЯ — имя первое, фамилия ЗАГЛАВНЫМИ",
  "title": "Полное название должности на русском",
  "company_name": "Название компании на языке оригинала",
  "company_description": "4-5 предложений: чем занимается, выручка/активы/сотрудники с конкретными цифрами, штаб-квартира (город страна), год основания, ключевые направления. Суммы: $ 4,6 млрд.",
  "company_intro": "1-2 предложения: с какого года и как компания сотрудничает с Узбекистаном или планирует. Конкретные даты если есть.",
  "company_uzbekistan": [
    { "sector": "Название отрасли", "description": "2-3 конкретных предложения: проекты, суммы, партнёры, договорённости. Используй данные из загруженного документа и дополнительного контекста в первую очередь." }
  ],
  "education": [
    { "years": "ГГГГ-ГГГГ гг.", "institution": "Название вуза", "country": "Страна", "degree": "Степень и специальность" }
  ],
  "career": [
    { "years": "ГГГГ-ГГГГ гг. или с ГГГГ г.", "company": "Название организации", "country": "Страна", "role": "Должность" }
  ]
}

Требования:
- company_uzbekistan: минимум 3 отрасли с конкретными деталями и цифрами
- education и career: только из реальных источников (LinkedIn, сайт, новости), хронологически
- career: минимум 4 записи включая текущую
- Годы: "1999-2007 гг." или "с 2021 г." — только дефис, не тире
- Суммы всегда с единицей: $ 110 млн., € 4,6 млрд.`;

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${OPENAI_API_KEY}` },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      max_tokens: 3000,
      response_format: { type: 'json_object' },
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  const data = await res.json();
  if (!data.choices?.[0]?.message?.content) throw new Error('OpenAI returned no content');
  const profile = JSON.parse(data.choices[0].message.content);
  return { ...profile, personImg, logoImg };
}

function validate(p) {
  for (const f of ['surname_caps','first_name','full_name_bio','title','company_name',
                    'company_description','company_uzbekistan','education','career']) {
    if (!p[f]) throw new Error(`Missing: ${f}`);
  }
  return p;
}

// ─── COMPANY SECTION (fits on 1 page each) ────────────────────────────────────
function buildCompanySection(profile) {
  const out = [];
  const cn = profile.company_name;

  // Heading
  out.push(cPara([
    r('ИНФОРМАЦИЯ', { bold: true, color: RED }),
    new TextRun({ break: 1 }),
    r('о компании ', { bold: true }),
    r('«', { bold: true }),
    r(cn, { bold: true, color: BLUE }),
    r('»', { bold: true }),
  ], { spacing: SP0 }));

  // Logo centered
  if (profile.logoImg) {
    const lp = imgPara(profile.logoImg, 5, 2);
    if (lp) out.push(lp);
  }
  out.push(ep());

  // Description — split into paragraphs
  const sents = profile.company_description.split(/(?<=[.!?])\s+/);
  let chunk = '';
  const chunks = [];
  for (const s of sents) {
    chunk += (chunk ? ' ' : '') + s;
    if (chunk.length > 300) { chunks.push(chunk); chunk = ''; }
  }
  if (chunk) chunks.push(chunk);
  for (const c of chunks) out.push(para(parseText(c, cn), { indent: true }));

  out.push(pb());

  // Uzbekistan heading
  out.push(cPara([
    r('Информация о деятельности компании', { bold: true }),
    new TextRun({ break: 1 }),
    r('«', { bold: true }),
    r(cn, { bold: true, color: BLUE }),
    r('» в Республике Узбекистан', { bold: true }),
  ], { spacing: SP0 }));
  out.push(ep());

  if (profile.company_intro) {
    out.push(para(parseText(profile.company_intro, cn), { indent: true }));
  }

  for (const s of profile.company_uzbekistan) {
    out.push(para([
      r(s.sector + ': ', { bold: true }),
      ...parseText(s.description, cn),
    ], { indent: true }));
  }

  return out;
}

// ─── PARTICIPANTS TABLE ───────────────────────────────────────────────────────
function formatDate(d, t) {
  const months = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
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
    new Paragraph({ children: [r(surname.toUpperCase(), { bold: true, color: BLUE }), ...(firstname ? [new TextRun({ break: 1 }), r(firstname, { color: BLUE })] : [])], spacing: SP0 }),
    epTiny(),
    ...(title ? [new Paragraph({ children: [r(title)], spacing: SP0, alignment: AlignmentType.BOTH })] : []),
  ], width);
}

function foreignCell(profile, width) {
  if (!profile) return tc([ep()], width);
  return tc([
    new Paragraph({ children: [r(profile.surname_caps, { bold: true, color: BLUE }), new TextRun({ break: 1 }), r(profile.first_name, { color: BLUE })], spacing: SP0 }),
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
    new TableRow({ children: [tc([para([r('Дата: ', { bold: true }), ...yearRuns(dateStr)])], CW, { span: 3 })] }),
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

// ─── BIO SECTION (transparent tables, centered photo) ────────────────────────
function buildBioSection(profile) {
  const out = [];
  out.push(pb());

  // Name — blue bold centered
  out.push(cPara([r(profile.full_name_bio, { bold: true, color: BLUE })], { spacing: SP0 }));

  // Photo centered (passport size 4×5cm)
  if (profile.personImg) {
    const pp = imgPara(profile.personImg, 4, 5);
    if (pp) out.push(pp);
  }
  out.push(ep());

  // Title + company — centered
  out.push(cPara([
    r(profile.title, { bold: true }),
    new TextRun({ break: 1 }),
    r('«'), r(profile.company_name, { bold: true, color: BLUE }), r('»'),
  ], { spacing: SP0 }));

  out.push(ep());

  const col1 = 2269;
  const col2 = 426;
  const col3 = CW - col1 - col2;

  // Education — NO BORDERS (transparent)
  const eduRows = [
    new TableRow({ children: [
      tc([para([r('Образование:', { bold: true })], { spacing: SP })], CW, { span: 3, noBorder: true }),
    ]}),
    ...profile.education.map(e => {
      const desc = [
        ...(e.degree ? [r(e.degree + ' ')] : []),
        ...(e.institution ? [r('«'), r(e.institution, { color: BLUE }), r('»')] : []),
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

  // Career — NO BORDERS (transparent)
  const careerRows = [
    new TableRow({ children: [
      tc([para([r('Профессиональная деятельность:', { bold: true })], { spacing: SP })], CW, { span: 3, noBorder: true }),
    ]}),
    ...profile.career.map(c => {
      const desc = [
        ...(c.role ? [r(c.role + ' ')] : []),
        ...(c.company ? [r('«'), r(c.company, { color: BLUE }), r('»')] : []),
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
  for (let i = 0; i < profiles.length; i++) {
    children = children.concat(buildCompanySection(profiles[i]));
    if (i < profiles.length - 1) children.push(pb());
  }
  children.push(pb());
  children = children.concat(buildParticipantsTable(date, time, uzbeks, profiles));
  for (const p of profiles) children = children.concat(buildBioSection(p));

  const doc = new Document({
    sections: [{ properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } }, children }],
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
  const { inputs, meeting_date, meeting_time, uzbek_participants, user_id, username, full_name, extra_context, extra_file } = req.body;

  if (!inputs?.length) return res.status(400).json({ error: 'inputs required' });
  if (!meeting_date) return res.status(400).json({ error: 'meeting_date required' });
  if (!uzbek_participants?.length) return res.status(400).json({ error: 'uzbek_participants required' });

  try {
    await upsertUser(user_id, username, full_name);

    // Inject extra context into each input
    const enrichedInputs = inputs.map(i => ({ ...i, extraContext: extra_context || null }));

    console.log(`Researching ${inputs.length} person(s)...`);
    const rawProfiles = await Promise.all(enrichedInputs.map(i => researchPerson(i)));
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

    const download_url = `${SUPABASE_URL}/storage/v1/object/public/documents/${fileName}`;

    await supabase.from('documents').insert({
      doc_id, user_id: String(user_id), download_url, meeting_date,
      profiles_count: profiles.length,
      queries: inputs.map(i => i.query),
      created_at: new Date().toISOString(),
    });

    console.log('Done:', doc_id);
    return res.json({ success: true, doc_id, download_url, meeting_date });
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
