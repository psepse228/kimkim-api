const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
        PageBreak } = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const TAVILY_API_KEY = process.env.TAVILY_API_KEY;
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ─── CONSTANTS (from template XML) ───────────────────────────────────────────
const FONT = 'Cambria';
const SZ = 30;          // 15pt
const SZ_SM = 6;        // tiny spacer
const RED = 'c00000';
const BLUE = '1f4e79';
const BLACK = '000000';
const SPACING = { after: 60, before: 60, line: 288, lineRule: 'auto' };
const FIRST_LINE = 709; // first line indent for body paragraphs

// A4 with ~2cm margins
const PAGE_W = 11906;
const MARGIN = 1134;
const CW = PAGE_W - MARGIN * 2; // ~9638 DXA

const BRD = { style: BorderStyle.SINGLE, size: 4, color: BLACK };
const BORDS = { top: BRD, bottom: BRD, left: BRD, right: BRD };

// ─── TEXT HELPERS ─────────────────────────────────────────────────────────────

function r(text, opts = {}) {
  return new TextRun({
    text: String(text),
    font: FONT,
    size: opts.sz || SZ,
    bold: opts.bold !== undefined ? opts.bold : false,
    italics: opts.italic || false,
    color: opts.color || BLACK,
  });
}

function rBold(text, color) { return r(text, { bold: true, color: color || BLACK }); }
function rRed(text) { return r(text, { bold: true, color: RED }); }
function rBlue(text) { return r(text, { bold: true, color: BLUE }); }
function rNormal(text) { return r(text, { bold: false, color: BLACK }); }

function para(runs, opts = {}) {
  return new Paragraph({
    children: Array.isArray(runs) ? runs : [runs],
    alignment: opts.align || AlignmentType.BOTH,
    spacing: opts.spacing || SPACING,
    indent: opts.indent ? { firstLine: FIRST_LINE } : undefined,
  });
}

function ep(sz) {
  return new Paragraph({
    children: [new TextRun({ text: '', font: FONT, size: sz || SZ })],
    spacing: SPACING,
  });
}

function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function tc(paragraphs, width, opts = {}) {
  return new TableCell({
    borders: opts.noBorder
      ? { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
      : BORDS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: opts.vAlign || VerticalAlign.CENTER,
    columnSpan: opts.span || 1,
    children: Array.isArray(paragraphs) ? paragraphs : [paragraphs],
  });
}

// ─── RESEARCH ────────────────────────────────────────────────────────────────

async function tavilySearch(query) {
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
  const snippets = (data.results || []).map(r => `${r.title}: ${r.content}`).join('\n\n');
  return (data.answer || '') + '\n\n' + snippets;
}

async function researchPerson(query) {
  const [personText, companyText] = await Promise.all([
    tavilySearch(`${query} biography career education background`),
    tavilySearch(`${query} company Uzbekistan cooperation investment projects`),
  ]);

  const prompt = `Ты — эксперт по составлению официальных справочных материалов для Министерства инвестиций Республики Узбекистан. Составь структурированные данные об иностранном бизнесмене для делового досье.

Данные о персоне:
${personText}

Данные о компании и Узбекистане:
${companyText}

Верни ТОЛЬКО JSON объект. Все поля на русском языке. Если информация не найдена — "Информация не найдена". Никогда null или пустые строки.

{
  "full_name": "ИМЯ ФАМИЛИЯ — имя с заглавной буквы, фамилия ЗАГЛАВНЫМИ (пример: РОБЕРТ ЛОВЕНТАЛЬ)",
  "full_name_reverse": "ФАМИЛИЯ Имя — фамилия ЗАГЛАВНЫМИ затем имя (пример: ЛОВЕНТАЛЬ Роберт)",
  "title": "Полное название должности на русском",
  "company_name": "Название компании на языке оригинала",
  "company_description": "4-5 предложений: чем занимается компания, масштаб (активы, сотрудники, офисы), штаб-квартира, год основания, ключевые направления. Упомянуть конкретные цифры если есть. Официальный стиль.",
  "company_intro": "1-2 предложения: когда и как компания начала сотрудничество с Узбекистаном. Конкретная дата/год если есть.",
  "company_uzbekistan": [
    {
      "sector": "Название отрасли",
      "description": "2-3 конкретных предложения: что именно делает/планирует компания в этой отрасли в Узбекистане. Суммы, проекты, партнёры если известны."
    }
  ],
  "education": [
    { "years": "ГГГГ-ГГГГ гг.", "description": "Степень, специальность, полное название вуза, страна." }
  ],
  "career": [
    { "years": "ГГГГ-ГГГГ гг. или с ГГГГ года", "description": "Должность в «Название организации»" }
  ]
}

Требования:
- full_name: формат "ИМЯ ФАМИЛИЯ" — имя первое, фамилия ЗАГЛАВНЫМИ (для заголовка биографии)
- full_name_reverse: формат "ФАМИЛИЯ Имя" — фамилия ЗАГЛАВНЫМИ, имя с заглавной (для таблицы участников)
- company_uzbekistan: минимум 3 отрасли, максимум 5, каждая с конкретными деталями
- education: хронологический порядок, минимум 1
- career: хронологический порядок, минимум 4 записи включая текущую
- Годы: "1999-2007 гг." или "с 2021 года"`;

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
  return JSON.parse(data.choices[0].message.content);
}

function validateProfile(p) {
  for (const f of ['full_name', 'full_name_reverse', 'title', 'company_name',
                    'company_description', 'company_uzbekistan', 'education', 'career']) {
    if (!p[f]) throw new Error(`Missing field: ${f}`);
  }
  if (!Array.isArray(p.company_uzbekistan) || p.company_uzbekistan.length < 1)
    throw new Error('company_uzbekistan empty');
  if (!Array.isArray(p.education) || p.education.length < 1)
    throw new Error('education empty');
  if (!Array.isArray(p.career) || p.career.length < 1)
    throw new Error('career empty');
  return p;
}

// ─── COMPANY SECTION ─────────────────────────────────────────────────────────

function buildCompanySection(profile) {
  const out = [];

  // Title: ИНФОРМАЦИЯ (red) + о деятельности (black) + «CompanyName» (blue) — centered
  out.push(para([
    rRed('ИНФОРМАЦИЯ '),
    rBold('\nо деятельности \n«'),
    rBlue(profile.company_name),
    rBold('»'),
  ], { align: AlignmentType.CENTER }));

  // Tiny spacer paragraph
  out.push(ep(SZ_SM));

  // Company description body paragraphs — justified with first line indent
  // Split on sentences for natural paragraph breaks
  const sentences = profile.company_description.split(/(?<=[.!?])\s+/);
  const mid = Math.ceil(sentences.length / 2);
  const para1 = sentences.slice(0, mid).join(' ');
  const para2 = sentences.slice(mid).join(' ');

  if (para1) out.push(para([rNormal(para1)], { indent: true }));
  if (para2) out.push(para([rNormal(para2)], { indent: true }));

  out.push(ep(SZ_SM));

  // Uzbekistan section title: ДЕЯТЕЛЬНОСТЬ (red) + company name (blue) + в Узбекистане (black)
  out.push(para([
    rRed('ДЕЯТЕЛЬНОСТЬ'),
    rBold('\n«'),
    rBlue(profile.company_name),
    rBold('» в Узбекистане'),
  ], { align: AlignmentType.CENTER }));

  out.push(ep(SZ_SM));

  // Intro paragraph about Uzbekistan cooperation
  if (profile.company_intro) {
    out.push(para([rNormal(profile.company_intro)], { indent: true }));
    out.push(ep(SZ_SM));
  }

  // Sector paragraphs: bold sector name + normal description
  for (const s of profile.company_uzbekistan) {
    out.push(para([
      rBold(`${s.sector}: `),
      rNormal(s.description),
    ], { indent: true }));
  }

  return out;
}

// ─── PARTICIPANTS TABLE ───────────────────────────────────────────────────────

function formatDate(meetingDate, meetingTime) {
  const months = ['января','февраля','марта','апреля','мая','июня',
                  'июля','августа','сентября','октября','ноября','декабря'];
  const d = new Date(meetingDate);
  return `${d.getDate()} ${months[d.getMonth()]} т.г., ${meetingTime}`;
}

function buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles) {
  const out = [];

  out.push(para([rBold('СПИСОК УЧАСТНИКОВ')], { align: AlignmentType.CENTER }));
  out.push(ep(SZ_SM));

  const dateStr = formatDate(meetingDate, meetingTime);
  const numW = 400;
  const colW = Math.floor((CW - numW) / 2);

  const rows = [];

  // Row 1: merged date cell
  rows.push(new TableRow({ children: [
    tc([para([rBold(`Дата: ${dateStr}`)])], CW, { span: 3 }),
  ]}));

  // Row 2: empty spacer
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([ep()], colW),
    tc([ep()], colW),
  ]}));

  // Row 3: headers
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([para([rBold('От узбекской стороны')], { align: AlignmentType.CENTER })], colW),
    tc([para([rBold('От иностранной стороны')], { align: AlignmentType.CENTER })], colW),
  ]}));

  // Rows 4-7: participants (always 4 rows)
  for (let i = 0; i < 4; i++) {
    const raw = uzbekParticipants[i] || '';
    const fp = profiles[i] || null;

    // Uzbek: parse "Name\nTitle" or just "Name"
    let uzbekParas;
    if (raw) {
      const [name, ...titleParts] = raw.split('\n');
      const title = titleParts.join(' ');
      uzbekParas = [
        para([rBlue(name.toUpperCase())]),
        title ? para([rNormal(title)]) : ep(SZ_SM),
      ];
    } else {
      uzbekParas = [ep()];
    }

    // Foreign: FULL_NAME_REVERSE (blue bold) + title (normal)
    let foreignParas;
    if (fp) {
      foreignParas = [
        para([rBlue(fp.full_name_reverse)]),
        para([rNormal(fp.title)]),
      ];
    } else {
      foreignParas = [ep()];
    }

    rows.push(new TableRow({ children: [
      tc([para([rBold(`${i + 1}.`)], { align: AlignmentType.CENTER })], numW),
      tc(uzbekParas, colW),
      tc(foreignParas, colW),
    ]}));
  }

  // Final empty row
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([ep()], colW),
    tc([ep()], colW),
  ]}));

  out.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [numW, colW, colW],
    rows,
  }));

  return out;
}

// ─── BIO SECTION ─────────────────────────────────────────────────────────────

function buildBioSection(profile) {
  const out = [];

  out.push(pb());

  // Name in blue bold (format: ИМЯ ФАМИЛИЯ)
  out.push(para([rBlue(profile.full_name)], { align: AlignmentType.CENTER }));

  // Title in black bold
  out.push(para([rBold(profile.title)], { align: AlignmentType.CENTER }));

  // Company in blue bold «CompanyName»
  out.push(para([
    rBold('«'),
    rBlue(profile.company_name),
    rBold('»'),
  ], { align: AlignmentType.CENTER }));

  out.push(ep(SZ_SM));

  const col1 = 2000; // years
  const col2 = 360;  // dash
  const col3 = CW - col1 - col2; // description

  // Education table
  const eduRows = [
    new TableRow({ children: [
      tc([para([rBold('Образование:')])], CW, { span: 3 }),
    ]}),
    ...profile.education.map(e => new TableRow({ children: [
      tc([para([rNormal(e.years)])], col1),
      tc([para([r('-', { italic: true })], { align: AlignmentType.CENTER })], col2),
      tc([para([rNormal(e.description)])], col3),
    ]})),
  ];

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: eduRows }));
  out.push(ep(SZ_SM));

  // Career table
  const careerRows = [
    new TableRow({ children: [
      tc([para([rBold('Профессиональная деятельность:')])], CW, { span: 3 }),
    ]}),
    ...profile.career.map(c => new TableRow({ children: [
      tc([para([rNormal(c.years)])], col1),
      tc([para([r('-', { italic: true })], { align: AlignmentType.CENTER })], col2),
      tc([para([rNormal(c.description)])], col3),
    ]})),
  ];

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: careerRows }));

  return out;
}

// ─── DOCX ASSEMBLY ────────────────────────────────────────────────────────────

async function generateDocx(profiles, meetingDate, meetingTime, uzbekParticipants) {
  let children = [];

  for (let i = 0; i < profiles.length; i++) {
    children = children.concat(buildCompanySection(profiles[i]));
    if (i < profiles.length - 1) children.push(pb());
  }

  children.push(pb());
  children = children.concat(buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles));

  for (const profile of profiles) {
    children = children.concat(buildBioSection(profile));
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

// ─── ENDPOINT ─────────────────────────────────────────────────────────────────

app.post('/generate', async (req, res) => {
  const { inputs, meeting_date, meeting_time, uzbek_participants, user_id } = req.body;

  if (!inputs?.length) return res.status(400).json({ error: 'inputs required' });
  if (!meeting_date) return res.status(400).json({ error: 'meeting_date required' });
  if (!uzbek_participants?.length) return res.status(400).json({ error: 'uzbek_participants required' });

  try {
    console.log(`Researching ${inputs.length} person(s)...`);
    const rawProfiles = await Promise.all(inputs.map(i => researchPerson(i.query)));
    const profiles = rawProfiles.map(validateProfile);
    console.log('Profiles validated.');

    console.log('Generating DOCX...');
    const buf = await generateDocx(profiles, meeting_date, meeting_time || '11:00', uzbek_participants);

    const doc_id = `doc_${Date.now()}_${user_id}`;
    const fileName = `${doc_id}.docx`;
    console.log(`Uploading ${fileName}...`);

    const { error: upErr } = await supabase.storage
      .from('documents')
      .upload(fileName, buf, {
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        upsert: true,
      });
    if (upErr) throw new Error(`Upload failed: ${upErr.message}`);

    const download_url = `${SUPABASE_URL}/storage/v1/object/public/documents/${fileName}`;

    const { error: dbErr } = await supabase.from('documents').insert({
      doc_id, user_id: String(user_id), download_url, meeting_date,
      profiles_count: profiles.length,
      queries: inputs.map(i => i.query),
      created_at: new Date().toISOString(),
    });
    if (dbErr) console.warn('DB warning:', dbErr.message);

    console.log('Done:', doc_id);
    return res.json({ success: true, doc_id, download_url, meeting_date });

  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/health', (_, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`KimKim API running on port ${PORT}`));
