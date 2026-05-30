const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak } = require('docx');
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

// ─── STYLE CONSTANTS (from template XML) ─────────────────────────────────────
const FONT = 'Cambria';
const SZ = 30;          // 15pt
const SZ_TINY = 10;     // 5pt spacer
const RED = 'C00000';
const BLUE = '1F4E79';
const BLACK = '000000';
const SP_BODY = { after: 120, before: 120, line: 288, lineRule: 'auto' };
const SP_COMPACT = { after: 0, line: 240, lineRule: 'auto' };
const SP_TINY = { after: 0, line: 20, lineRule: 'atLeast' };
const INDENT_FIRST = { firstLine: 720 }; // 1.25cm

// A4 with 2cm margins
const PAGE_W = 11906;
const MARGIN = 1134;
const CW = PAGE_W - MARGIN * 2; // 9638 DXA

const BRD = { style: BorderStyle.SINGLE, size: 4, color: BLACK };
const BORDS = { top: BRD, bottom: BRD, left: BRD, right: BRD };
const NO_BORD = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDS = { top: NO_BORD, bottom: NO_BORD, left: NO_BORD, right: NO_BORD };

// ─── TEXT RUN HELPERS ────────────────────────────────────────────────────────

function run(text, opts = {}) {
  return new TextRun({
    text: String(text),
    font: FONT,
    size: opts.sz || SZ,
    bold: opts.bold || false,
    italics: opts.italic || false,
    color: opts.color || undefined,
    smallCaps: opts.smallCaps || false,
  });
}

// Parse a year string like "1989-1993 гг." into colored runs:
// digits = red, separators/text = black
function yearRuns(yearStr) {
  // Split into tokens: digit sequences vs non-digit
  const tokens = yearStr.split(/(\d+)/);
  return tokens.filter(t => t.length > 0).map(t => {
    if (/^\d+$/.test(t)) return run(t, { color: RED });
    return run(t);
  });
}

// Parse description: detect institution names (wrapped in « »), country (parentheses), years
function descRuns(text) {
  if (!text) return [run('')];
  const runs = [];
  // Tokenize: «...» = blue, (Страна) = italic blue for geo, digits = red
  const regex = /«([^»]+)»|\(([^)]+)\)|(\d+)/g;
  let last = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > last) runs.push(run(text.slice(last, match.index)));
    if (match[1] !== undefined) {
      // Company name «...»
      runs.push(run('«'));
      runs.push(run(match[1], { color: BLUE, bold: true }));
      runs.push(run('»'));
    } else if (match[2] !== undefined) {
      // Parenthetical: (content) — check if geo or year
      const inner = match[2];
      if (/^\d/.test(inner)) {
        // Year reference
        runs.push(run('(', { italic: true }));
        runs.push(run(inner, { italic: true, color: RED }));
        runs.push(run(')', { italic: true }));
      } else {
        // Geographic/org reference
        runs.push(run('(', { italic: true }));
        runs.push(run(inner, { italic: true, color: BLUE }));
        runs.push(run(')', { italic: true }));
      }
    } else if (match[3] !== undefined) {
      // Standalone digits
      runs.push(run(match[3], { color: RED }));
    }
    last = match.index + match[0].length;
  }
  if (last < text.length) runs.push(run(text.slice(last)));
  return runs.length > 0 ? runs : [run(text)];
}

function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align || AlignmentType.BOTH,
    spacing: opts.spacing || SP_BODY,
    indent: opts.indent ? INDENT_FIRST : undefined,
  });
}

function centerPara(children, opts = {}) {
  return para(children, { ...opts, align: AlignmentType.CENTER });
}

function ep(sz) {
  return new Paragraph({
    children: [new TextRun({ text: '', font: FONT, size: sz || SZ })],
    spacing: SP_COMPACT,
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

  const prompt = `Ты — эксперт по составлению официальных справочных материалов для Министерства инвестиций Республики Узбекистан. Составь структурированные данные об иностранном бизнесмене для делового досье к встрече.

Данные о персоне:
${personText}

Данные о компании и Узбекистане:
${companyText}

Верни ТОЛЬКО JSON объект. Все поля СТРОГО на русском языке. Если информация не найдена — "Информация не найдена". Никогда null или пустые строки.

{
  "surname_caps": "ФАМИЛИЯ заглавными буквами (только фамилия, пример: РОУ)",
  "first_name": "Имя с заглавной буквы (только имя, пример: Скотт)",
  "full_name_bio": "ИМЯ ФАМИЛИЯ — имя первое, фамилия ЗАГЛАВНЫМИ (для заголовка биографии, пример: СКОТТ РОУ)",
  "title": "Полное название должности на русском",
  "company_name": "Название компании на языке оригинала (латиница)",
  "company_description": "4-5 предложений о компании. Включить: чем занимается, масштаб (выручка/активы/сотрудники с конкретными цифрами), штаб-квартира (город, страна), год основания, ключевые продукты/услуги. Официальный деловой стиль. Компании и города писать в оригинале или в скобках как принято. Пример формата сумм: $ 4,6 млрд.",
  "company_intro": "1-2 предложения: с какого года и как компания сотрудничает с Узбекистаном или планирует сотрудничество.",
  "company_uzbekistan": [
    {
      "sector": "Название отрасли (конкретное, не общее)",
      "description": "2-3 предложения: конкретные проекты, суммы, партнёры, договорённости в этой отрасли в Узбекистане."
    }
  ],
  "education": [
    {
      "years": "ГГГГ-ГГГГ гг. или ГГГГ г.",
      "institution": "Полное название учебного заведения",
      "country": "Страна",
      "degree": "Степень и специальность"
    }
  ],
  "career": [
    {
      "years": "ГГГГ-ГГГГ гг. или с ГГГГ г.",
      "company": "Название организации",
      "country": "Страна (если известна)",
      "role": "Должность"
    }
  ]
}

Требования:
- company_uzbekistan: минимум 3 отрасли, каждая с конкретными деталями
- education: хронологически, минимум 1
- career: хронологически, минимум 4 записи
- Годы только цифры: "1999-2007 гг." или "с 2021 г."`;

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
  for (const f of ['surname_caps', 'first_name', 'full_name_bio', 'title', 'company_name',
                    'company_description', 'company_uzbekistan', 'education', 'career']) {
    if (!p[f]) throw new Error(`Missing field: ${f}`);
  }
  return p;
}

// ─── COMPANY SECTION ─────────────────────────────────────────────────────────

function buildCompanySection(profile) {
  const out = [];

  // Heading: ИНФОРМАЦИЯ (red, bold) line break, о компании (black bold), «CompanyName» (blue bold)
  out.push(centerPara([
    run('ИНФОРМАЦИЯ', { bold: true, color: RED }),
    new TextRun({ break: 1 }),
    run('о компании ', { bold: true }),
    run('«', { bold: true }),
    run(profile.company_name, { bold: true, color: BLUE }),
    run('»', { bold: true }),
  ], { spacing: SP_COMPACT }));

  out.push(ep());

  // Company description — justified with first-line indent
  // Split into multiple paragraphs for readability
  const sentences = profile.company_description.split(/(?<=[.!?])\s+/);
  const chunks = [];
  let chunk = '';
  for (const s of sentences) {
    chunk += (chunk ? ' ' : '') + s;
    if (chunk.length > 200) { chunks.push(chunk); chunk = ''; }
  }
  if (chunk) chunks.push(chunk);

  for (const chunk of chunks) {
    out.push(para(descRuns(chunk), { indent: true }));
  }

  out.push(ep());

  // Uzbekistan section heading
  out.push(centerPara([
    run('Информация о деятельности компании', { bold: true }),
    new TextRun({ break: 1 }),
    run('«', { bold: true }),
    run(profile.company_name, { bold: true, color: BLUE }),
    run('» в Республике Узбекистан', { bold: true }),
  ], { spacing: SP_COMPACT }));

  out.push(ep());

  // Intro
  if (profile.company_intro) {
    out.push(para(descRuns(profile.company_intro), { indent: true }));
  }

  // Sector paragraphs
  for (const s of profile.company_uzbekistan) {
    const sectorRuns = [
      run(s.sector + ': ', { bold: true }),
      ...descRuns(s.description),
    ];
    out.push(para(sectorRuns, { indent: true }));
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

// Participant cell for uzbek side: SURNAME (blue bold) + break + first name (blue not bold) + spacer para + title (black)
function uzbekCell(nameRaw, width) {
  if (!nameRaw) return tc([ep()], width);
  const parts = nameRaw.split('\n');
  const namePart = parts[0] || '';
  const titlePart = parts.slice(1).join(' ') || '';

  // Try to split name into surname + firstname
  const words = namePart.trim().split(/\s+/);
  const surname = words[0] || '';
  const firstname = words.slice(1).join(' ');

  const paras = [];
  // Para 1: SURNAME (blue bold) + break + firstname (blue, not bold)
  const nameRunChildren = [run(surname.toUpperCase(), { bold: true, color: BLUE })];
  if (firstname) {
    nameRunChildren.push(new TextRun({ break: 1 }));
    nameRunChildren.push(run(firstname, { color: BLUE, bold: false }));
  }
  paras.push(new Paragraph({ children: nameRunChildren, spacing: SP_COMPACT }));

  // Tiny spacer
  paras.push(epTiny());

  // Title
  if (titlePart) {
    paras.push(new Paragraph({
      children: [run(titlePart)],
      spacing: SP_COMPACT,
      alignment: AlignmentType.BOTH,
    }));
  }

  return tc(paras, width);
}

// Participant cell for foreign side: SURNAME (blue bold) + break + firstname (blue, not bold) + spacer + title (black)
function foreignCell(profile, width) {
  if (!profile) return tc([ep()], width);

  const paras = [];
  paras.push(new Paragraph({
    children: [
      run(profile.surname_caps, { bold: true, color: BLUE }),
      new TextRun({ break: 1 }),
      run(profile.first_name, { color: BLUE, bold: false }),
    ],
    spacing: SP_COMPACT,
  }));

  paras.push(epTiny());

  paras.push(new Paragraph({
    children: [run(profile.title)],
    spacing: SP_COMPACT,
    alignment: AlignmentType.BOTH,
  }));

  return tc(paras, width);
}

function buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles) {
  const out = [];

  out.push(centerPara([run('СПИСОК УЧАСТНИКОВ', { bold: true })], { spacing: SP_COMPACT }));
  out.push(ep());

  const dateStr = formatDate(meetingDate, meetingTime);
  const numW = 400;
  const colW = Math.floor((CW - numW) / 2);

  const rows = [];

  // Date header row - merged across all columns
  rows.push(new TableRow({ children: [
    tc([new Paragraph({
      children: [run('Дата: ', { bold: true }), ...yearRuns(dateStr)],
      spacing: SP_COMPACT,
    })], CW, { span: 3 }),
  ]}));

  // Empty spacer row
  rows.push(new TableRow({ children: [
    tc([ep()], numW), tc([ep()], colW), tc([ep()], colW),
  ]}));

  // Header row
  rows.push(new TableRow({ children: [
    tc([ep()], numW),
    tc([centerPara([run('От узбекской стороны', { bold: true })], { spacing: SP_COMPACT })], colW),
    tc([centerPara([run('От иностранной стороны', { bold: true })], { spacing: SP_COMPACT })], colW),
  ]}));

  // 4 participant rows
  const maxRows = 4;
  for (let i = 0; i < maxRows; i++) {
    rows.push(new TableRow({ children: [
      tc([centerPara([run(`${i + 1}.`, { bold: true })], { spacing: SP_COMPACT })], numW),
      uzbekCell(uzbekParticipants[i] || '', colW),
      foreignCell(profiles[i] || null, colW),
    ]}));
  }

  // Final empty row
  rows.push(new TableRow({ children: [
    tc([ep()], numW), tc([ep()], colW), tc([ep()], colW),
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

  // Name heading: full_name_bio (blue, bold, smallCaps style)
  out.push(centerPara([
    run(profile.full_name_bio, { bold: true, color: BLUE }),
  ], { spacing: SP_COMPACT }));

  // Title + company (same paragraph, line break between)
  out.push(centerPara([
    run(profile.title, { bold: true }),
    new TextRun({ break: 1 }),
    run('«', { bold: false }),
    run(profile.company_name, { bold: true, color: BLUE }),
    run('»', { bold: false }),
  ], { spacing: SP_COMPACT }));

  out.push(ep());

  // Column widths from template XML: 2269, 426, rest
  const col1 = 2269;
  const col2 = 426;
  const col3 = CW - col1 - col2;

  // Education table
  const eduRows = [
    new TableRow({ children: [
      tc([new Paragraph({
        children: [run('Образование:', { bold: true })],
        spacing: SP_BODY,
      })], CW, { span: 3 }),
    ]}),
    ...profile.education.map(e => {
      // Years: digits red, text black
      const yearsRuns = yearRuns(e.years);

      // Description: degree + institution (blue) + (country) italic blue
      const descChildren = [];
      if (e.degree) descChildren.push(run(e.degree + ' '));
      if (e.institution) {
        descChildren.push(run('«'));
        descChildren.push(run(e.institution, { color: BLUE }));
        descChildren.push(run('»'));
      }
      if (e.country) {
        descChildren.push(run(' (', { italic: true }));
        descChildren.push(run(e.country, { italic: true, color: BLUE }));
        descChildren.push(run(')', { italic: true }));
      }

      return new TableRow({ children: [
        tc([new Paragraph({ children: yearsRuns, spacing: SP_BODY })], col1),
        tc([new Paragraph({ children: [run('-', { italic: true })], spacing: SP_BODY, alignment: AlignmentType.CENTER })], col2),
        tc([new Paragraph({ children: descChildren.length ? descChildren : [run(e.degree || 'Информация не найдена')], spacing: SP_BODY })], col3),
      ]});
    }),
  ];

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: eduRows }));
  out.push(ep());

  // Career table
  const careerRows = [
    new TableRow({ children: [
      tc([new Paragraph({
        children: [run('Профессиональная деятельность:', { bold: true })],
        spacing: SP_BODY,
      })], CW, { span: 3 }),
    ]}),
    ...profile.career.map(c => {
      const yearsRuns = yearRuns(c.years);

      const descChildren = [];
      if (c.role) descChildren.push(run(c.role + ' '));
      if (c.company) {
        descChildren.push(run('«'));
        descChildren.push(run(c.company, { color: BLUE }));
        descChildren.push(run('»'));
      }
      if (c.country) {
        descChildren.push(run(' (', { italic: true }));
        descChildren.push(run(c.country, { italic: true, color: BLUE }));
        descChildren.push(run(')', { italic: true }));
      }

      return new TableRow({ children: [
        tc([new Paragraph({ children: yearsRuns, spacing: SP_BODY })], col1),
        tc([new Paragraph({ children: [run('-', { italic: true })], spacing: SP_BODY, alignment: AlignmentType.CENTER })], col2),
        tc([new Paragraph({ children: descChildren.length ? descChildren : [run(c.role || '')], spacing: SP_BODY })], col3),
      ]});
    }),
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
