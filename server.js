const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
        PageBreak } = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
app.use(express.json());
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
      include_answer: true
    })
  });
  const data = await res.json();
  const snippets = (data.results || []).map(r => `${r.title}: ${r.content}`).join('\n\n');
  return (data.answer || '') + '\n\n' + snippets;
}

async function researchPerson(query) {
  const searchText = await tavilySearch(`${query} biography career education background`);
  const companyQuery = query.includes('http') ? query : `${query} company`;
  const companyText = await tavilySearch(`${companyQuery} Uzbekistan cooperation investment projects`);

  const prompt = `Ты — эксперт по составлению справочных материалов для Министерства инвестиций Узбекистана. Тебе нужно извлечь структурированные данные об иностранном бизнесмене для официального документа к деловой встрече.

Данные о персоне:
${searchText}

Данные о компании и связи с Узбекистаном:
${companyText}

Верни ТОЛЬКО JSON объект со следующей точной структурой. Все текстовые поля должны быть на русском языке. Если информация не найдена — используй "Информация не найдена". Никогда не возвращай null или пустые строки.

{
  "full_name": "ФАМИЛИЯ Имя — фамилия заглавными буквами, затем имя с заглавной буквы",
  "title": "Полное название должности в компании на русском языке",
  "company_name": "Название компании на языке оригинала",
  "company_description": "4-5 предложений о компании: чем занимается, активы/масштаб, штаб-квартира, история, ключевые направления деятельности. Профессиональный деловой стиль. На русском.",
  "company_uzbekistan": [
    {
      "sector": "Название отрасли на русском",
      "description": "2-3 предложения о том, что компания делает или планирует делать в данной отрасли в Узбекистане. Конкретные факты: суммы инвестиций, проекты, партнёры, договорённости."
    }
  ],
  "education": [
    {
      "years": "ГГГГ-ГГГГ гг.",
      "description": "Степень и специальность, полное название учебного заведения, страна. На русском."
    }
  ],
  "career": [
    {
      "years": "ГГГГ-ГГГГ гг. или с ГГГГ года",
      "description": "Должность и полное название организации. На русском."
    }
  ]
}

Строгие требования:
- full_name: фамилия ВСЯ ЗАГЛАВНЫМИ, затем имя с заглавной буквы (пример: ИВАНОВ Александр)
- company_description: минимум 4 предложения, включая ключевые финансовые показатели если есть
- company_uzbekistan: минимум 3 отрасли, максимум 5. Каждая — конкретная, с деталями. Не общие фразы.
- education: хронологический порядок, минимум 1 запись
- career: хронологический порядок, минимум 4 записи, включая текущую должность
- Все годы в русском формате: "1999-2007 гг." или "с 2021 года"
- Стиль — официальный деловой документ Министерства`;

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${OPENAI_API_KEY}`
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      max_tokens: 3000,
      response_format: { type: 'json_object' },
      messages: [{ role: 'user', content: prompt }]
    })
  });

  const data = await res.json();
  if (!data.choices?.[0]?.message?.content) {
    throw new Error('OpenAI returned no content');
  }
  return JSON.parse(data.choices[0].message.content);
}

// ─── VALIDATION ───────────────────────────────────────────────────────────────

function validateProfile(profile) {
  const required = ['full_name', 'title', 'company_name', 'company_description',
                    'company_uzbekistan', 'education', 'career'];
  for (const field of required) {
    if (!profile[field]) throw new Error(`Missing field: ${field}`);
  }
  if (!Array.isArray(profile.company_uzbekistan) || profile.company_uzbekistan.length < 1)
    throw new Error('company_uzbekistan must have at least 1 entry');
  if (!Array.isArray(profile.education) || profile.education.length < 1)
    throw new Error('education must have at least 1 entry');
  if (!Array.isArray(profile.career) || profile.career.length < 1)
    throw new Error('career must have at least 1 entry');
  return profile;
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

const FONT = 'Times New Roman';
const PAGE_WIDTH = 11906;
const MARGIN = 1134; // ~2cm margins
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // ~9638

const brd = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDERS = { top: brd, bottom: brd, left: brd, right: brd };
const NO_BRD = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BRD, bottom: NO_BRD, left: NO_BRD, right: NO_BRD };

function t(text, opts = {}) {
  return new TextRun({ text: String(text), font: FONT, size: 24, bold: false, italics: false, ...opts });
}

function p(runs, opts = {}) {
  return new Paragraph({
    children: Array.isArray(runs) ? runs : [runs],
    spacing: { after: opts.after !== undefined ? opts.after : 100, before: opts.before || 0, line: opts.line || 276 },
    alignment: opts.align || AlignmentType.BOTH,
  });
}

function ep() { return p([t('')], { after: 60 }); }
function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function tc(paragraphs, width, opts = {}) {
  return new TableCell({
    borders: opts.noBorder ? NO_BORDERS : BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: opts.vAlign || VerticalAlign.CENTER,
    columnSpan: opts.span || 1,
    children: Array.isArray(paragraphs) ? paragraphs : [paragraphs]
  });
}

function formatDate(meetingDate, meetingTime) {
  const months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
                  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const d = new Date(meetingDate);
  const day = d.getDate();
  const month = months[d.getMonth()];
  return `${day} ${month} т.г., ${meetingTime}`;
}

// ─── COMPANY SECTION ─────────────────────────────────────────────────────────

function buildCompanySection(profile) {
  const children = [];

  children.push(p([
    t('ИНФОРМАЦИЯ ', { bold: true }),
    t('о деятельности ', { bold: true }),
    t(`«${profile.company_name}»`, { bold: true })
  ], { after: 160, align: AlignmentType.CENTER }));

  children.push(p([t(profile.company_description)], { after: 200 }));

  children.push(p([
    t('ДЕЯТЕЛЬНОСТЬ', { bold: true }),
    t(` «${profile.company_name}»`, { bold: true }),
    t(' в Узбекистане', { bold: true })
  ], { after: 160, align: AlignmentType.CENTER }));

  for (const sector of profile.company_uzbekistan) {
    children.push(p([
      t(`${sector.sector}: `, { bold: true }),
      t(sector.description)
    ], { after: 140 }));
  }

  return children;
}

// ─── PARTICIPANTS TABLE ───────────────────────────────────────────────────────

function buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles) {
  const children = [];

  children.push(p([t('СПИСОК УЧАСТНИКОВ', { bold: true })],
    { after: 140, align: AlignmentType.CENTER }));

  const dateStr = formatDate(meetingDate, meetingTime);
  const numW = 380;
  const colW = Math.floor((CONTENT_WIDTH - numW) / 2);

  const rows = [];

  // Row 1: date merged across all 3 columns
  rows.push(new TableRow({ children: [
    tc([p([t(`Дата: ${dateStr}`, { bold: true })], { after: 40, align: AlignmentType.LEFT })],
      CONTENT_WIDTH, { span: 3 })
  ]}));

  // Row 2: empty spacer
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([ep()], colW),
    tc([ep()], colW)
  ]}));

  // Row 3: column headers
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([p([t('От узбекской стороны', { bold: true })], { after: 40, align: AlignmentType.CENTER })], colW),
    tc([p([t('От иностранной стороны', { bold: true })], { after: 40, align: AlignmentType.CENTER })], colW)
  ]}));

  // Rows 4-7: participants (max 4)
  for (let i = 0; i < 4; i++) {
    const uzbekRaw = uzbekParticipants[i] || '';
    const foreignProfile = profiles[i] || null;

    // Parse uzbek participant: "ФАМИЛИЯ Имя\nДолжность" or just name
    let uzbekParagraphs;
    if (uzbekRaw) {
      const parts = uzbekRaw.split('\n');
      const namePart = parts[0] || '';
      const titlePart = parts[1] || '';
      uzbekParagraphs = [
        p([t(namePart.toUpperCase(), { bold: true })], { after: 20 }),
        titlePart ? p([t(titlePart)], { after: 40 }) : ep()
      ];
    } else {
      uzbekParagraphs = [ep()];
    }

    let foreignParagraphs;
    if (foreignProfile) {
      foreignParagraphs = [
        p([t(foreignProfile.full_name, { bold: true })], { after: 20 }),
        p([t(foreignProfile.title)], { after: 40 })
      ];
    } else {
      foreignParagraphs = [ep()];
    }

    rows.push(new TableRow({ children: [
      tc([p([t(`${i + 1}.`, { bold: true })], { after: 40, align: AlignmentType.CENTER })], numW),
      tc(uzbekParagraphs, colW),
      tc(foreignParagraphs, colW)
    ]}));
  }

  // Empty last row
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([ep()], colW),
    tc([ep()], colW)
  ]}));

  children.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [numW, colW, colW],
    rows
  }));

  return children;
}

// ─── BIO SECTION ─────────────────────────────────────────────────────────────

function buildBioSection(profile) {
  const children = [];

  children.push(pb());

  children.push(p([t(profile.full_name, { bold: true })],
    { after: 40, align: AlignmentType.CENTER }));
  children.push(p([t(profile.title, { bold: true })],
    { after: 40, align: AlignmentType.CENTER }));

  // Company name under title
  children.push(p([t(`«${profile.company_name}»`, { bold: true })],
    { after: 200, align: AlignmentType.CENTER }));

  const col1 = 1900;
  const col2 = 380;
  const col3 = CONTENT_WIDTH - col1 - col2;

  // Education table
  const eduRows = [
    new TableRow({ children: [
      tc([p([t('Образование:', { bold: true })], { after: 40 })],
        CONTENT_WIDTH, { span: 3 })
    ]}),
    ...profile.education.map(e => new TableRow({ children: [
      tc([p([t(e.years)], { after: 40 })], col1),
      tc([p([t('-', { italics: true })], { after: 40, align: AlignmentType.CENTER })], col2),
      tc([p([t(e.description)], { after: 40 })], col3)
    ]}))
  ];

  children.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col1, col2, col3],
    rows: eduRows
  }));

  children.push(ep());

  // Career table
  const careerRows = [
    new TableRow({ children: [
      tc([p([t('Профессиональная деятельность:', { bold: true })], { after: 40 })],
        CONTENT_WIDTH, { span: 3 })
    ]}),
    ...profile.career.map(c => new TableRow({ children: [
      tc([p([t(c.years)], { after: 40 })], col1),
      tc([p([t('-', { italics: true })], { after: 40, align: AlignmentType.CENTER })], col2),
      tc([p([t(c.description)], { after: 40 })], col3)
    ]}))
  ];

  children.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col1, col2, col3],
    rows: careerRows
  }));

  return children;
}

// ─── DOCX ASSEMBLY ────────────────────────────────────────────────────────────

async function generateDocx(profiles, meetingDate, meetingTime, uzbekParticipants) {
  let allChildren = [];

  for (let i = 0; i < profiles.length; i++) {
    allChildren = allChildren.concat(buildCompanySection(profiles[i]));
    if (i < profiles.length - 1) allChildren.push(pb());
  }

  allChildren.push(pb());
  allChildren = allChildren.concat(
    buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles)
  );

  for (const profile of profiles) {
    allChildren = allChildren.concat(buildBioSection(profile));
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
        }
      },
      children: allChildren
    }]
  });

  return await Packer.toBuffer(doc);
}

// ─── MAIN ENDPOINT ────────────────────────────────────────────────────────────

app.post('/generate', async (req, res) => {
  const { inputs, meeting_date, meeting_time, uzbek_participants, user_id } = req.body;

  if (!inputs || !inputs.length) return res.status(400).json({ error: 'inputs required' });
  if (!meeting_date) return res.status(400).json({ error: 'meeting_date required' });
  if (!uzbek_participants || !uzbek_participants.length)
    return res.status(400).json({ error: 'uzbek_participants required' });

  try {
    console.log(`Researching ${inputs.length} person(s)...`);
    const rawProfiles = await Promise.all(inputs.map(i => researchPerson(i.query)));
    const profiles = rawProfiles.map(validateProfile);
    console.log('All profiles validated.');

    console.log('Generating DOCX...');
    const docBuffer = await generateDocx(
      profiles, meeting_date, meeting_time || '11:00', uzbek_participants
    );

    const doc_id = `doc_${Date.now()}_${user_id}`;
    const fileName = `${doc_id}.docx`;
    console.log(`Uploading ${fileName}...`);

    const { error: uploadError } = await supabase.storage
      .from('documents')
      .upload(fileName, docBuffer, {
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        upsert: true
      });

    if (uploadError) throw new Error(`Supabase upload failed: ${uploadError.message}`);

    const download_url = `${SUPABASE_URL}/storage/v1/object/public/documents/${fileName}`;

    const { error: dbError } = await supabase
      .from('documents')
      .insert({
        doc_id,
        user_id: String(user_id),
        download_url,
        meeting_date,
        profiles_count: profiles.length,
        queries: inputs.map(i => i.query),
        created_at: new Date().toISOString()
      });

    if (dbError) console.warn('DB insert warning:', dbError.message);

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
