const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak,
        ImageRun } = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
app.use(express.json({ limit: '10mb' }));
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
const SZ = 30;
const SZ_TINY = 10;
const RED = 'C00000';
const BLUE = '1F4E79';
const BLACK = '000000';
const SP_BODY = { after: 120, before: 120, line: 288, lineRule: 'auto' };
const SP_COMPACT = { after: 0, line: 240, lineRule: 'auto' };
const SP_TINY = { after: 0, line: 20, lineRule: 'atLeast' };
const INDENT_FIRST = { firstLine: 720 };
const PAGE_W = 11906;
const MARGIN = 1134;
const CW = PAGE_W - MARGIN * 2;
const BRD = { style: BorderStyle.SINGLE, size: 4, color: BLACK };
const BORDS = { top: BRD, bottom: BRD, left: BRD, right: BRD };
const NO_BORD = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDS = { top: NO_BORD, bottom: NO_BORD, left: NO_BORD, right: NO_BORD };

// ─── TEXT HELPERS ─────────────────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({
    text: String(text),
    font: FONT,
    size: opts.sz || SZ,
    bold: opts.bold || false,
    italics: opts.italic || false,
    color: opts.color || undefined,
  });
}

function yearRuns(yearStr) {
  const tokens = yearStr.split(/(\d+)/);
  return tokens.filter(t => t.length > 0).map(t =>
    /^\d+$/.test(t) ? run(t, { color: RED }) : run(t)
  );
}

function descRuns(text, companyName) {
  if (!text) return [run('')];
  const runs = [];
  // Match «...», (geo/year), digits, and company name plain mentions
  const escapedCompany = companyName ? companyName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') : null;
  const pattern = escapedCompany
    ? `«([^»]+)»|\\(([^)]+)\\)|(${escapedCompany})|(\\d+)`
    : `«([^»]+)»|\\(([^)]+)\\)|(\\d+)`;
  const regex = new RegExp(pattern, 'g');
  let last = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > last) runs.push(run(text.slice(last, match.index)));
    if (match[1] !== undefined) {
      runs.push(run('«'));
      runs.push(run(match[1], { color: BLUE, bold: true }));
      runs.push(run('»'));
    } else if (match[2] !== undefined) {
      const inner = match[2];
      if (/^\d/.test(inner)) {
        runs.push(run('(', { italic: true }));
        runs.push(run(inner, { italic: true, color: RED }));
        runs.push(run(')', { italic: true }));
      } else {
        runs.push(run('(', { italic: true }));
        runs.push(run(inner, { italic: true, color: BLUE }));
        runs.push(run(')', { italic: true }));
      }
    } else if (escapedCompany && match[3] !== undefined) {
      runs.push(run(match[3], { color: BLUE, bold: true }));
    } else {
      const digitMatch = escapedCompany ? match[4] : match[3];
      if (digitMatch) runs.push(run(digitMatch, { color: RED }));
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

// ─── IMAGE FETCHING ───────────────────────────────────────────────────────────
async function fetchImageBuffer(url) {
  try {
    const res = await fetch(url, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      signal: AbortSignal.timeout(8000),
    });
    if (!res.ok) return null;
    const contentType = res.headers.get('content-type') || '';
    if (!contentType.includes('image')) return null;
    const arrayBuffer = await res.arrayBuffer();
    return { buffer: Buffer.from(arrayBuffer), contentType };
  } catch {
    return null;
  }
}

function getImageType(contentType) {
  if (contentType.includes('png')) return 'png';
  if (contentType.includes('gif')) return 'gif';
  if (contentType.includes('webp')) return 'png'; // fallback
  return 'jpg';
}

async function searchImage(query) {
  try {
    const res = await fetch('https://api.tavily.com/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        api_key: TAVILY_API_KEY,
        query,
        search_depth: 'basic',
        max_results: 5,
        include_images: true,
      }),
    });
    const data = await res.json();
    const images = data.images || [];
    for (const imgUrl of images) {
      if (!imgUrl || imgUrl.includes('svg')) continue;
      const result = await fetchImageBuffer(imgUrl);
      if (result) return result;
    }
    return null;
  } catch {
    return null;
  }
}

function makeImagePara(imgData, widthCm, heightCm) {
  if (!imgData) return null;
  // Convert cm to EMU (1cm = 360000 EMU)
  const w = Math.round(widthCm * 360000);
  const h = Math.round(heightCm * 360000);
  try {
    return centerPara([
      new ImageRun({
        data: imgData.buffer,
        transformation: { width: Math.round(widthCm * 28.35), height: Math.round(heightCm * 28.35) },
        type: getImageType(imgData.contentType),
      }),
    ], { spacing: SP_COMPACT });
  } catch {
    return null;
  }
}

// ─── RESEARCH ─────────────────────────────────────────────────────────────────
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
  const [personText, companyText, personImg, logoImg] = await Promise.all([
    tavilySearch(`${query} biography career education background`),
    tavilySearch(`${query} company Uzbekistan cooperation investment projects`),
    searchImage(`${query} portrait photo professional`),
    null, // logo fetched per company below
  ]);

  const prompt = `Ты — эксперт по составлению официальных справочных материалов для Министерства инвестиций Республики Узбекистан.

Данные о персоне:
${personText}

Данные о компании и Узбекистане:
${companyText}

Верни ТОЛЬКО JSON. Все поля на русском языке. Если информация не найдена — "Информация не найдена". Никогда null или пустые строки.

{
  "surname_caps": "ФАМИЛИЯ заглавными (только фамилия)",
  "first_name": "Имя с заглавной буквы",
  "full_name_bio": "ИМЯ ФАМИЛИЯ — имя первое, фамилия ЗАГЛАВНЫМИ (пример: СКОТТ РОУ)",
  "title": "Полное название должности на русском",
  "company_name": "Название компании на языке оригинала (латиница)",
  "company_description": "4-5 предложений: чем занимается, выручка/активы/сотрудники с цифрами, штаб-квартира (город, страна), год основания, ключевые продукты. Суммы в формате: $ 4,6 млрд.",
  "company_intro": "1-2 предложения: с какого года и как компания сотрудничает с Узбекистаном.",
  "company_uzbekistan": [
    { "sector": "Название отрасли", "description": "2-3 предложения с конкретными проектами, суммами, партнёрами." }
  ],
  "education": [
    { "years": "ГГГГ-ГГГГ гг.", "institution": "Название вуза", "country": "Страна", "degree": "Степень и специальность" }
  ],
  "career": [
    { "years": "ГГГГ-ГГГГ гг. или с ГГГГ г.", "company": "Название организации", "country": "Страна", "role": "Должность" }
  ]
}

Требования:
- company_uzbekistan: минимум 3 отрасли с конкретными деталями
- education: хронологически, минимум 1
- career: хронологически, минимум 4 записи
- Годы: "1999-2007 гг." или "с 2021 г."`;

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

  // Fetch company logo
  const logo = await searchImage(`${profile.company_name} company logo official`);

  return { ...profile, personImg, logoImg: logo };
}

function validateProfile(p) {
  for (const f of ['surname_caps', 'first_name', 'full_name_bio', 'title', 'company_name',
                    'company_description', 'company_uzbekistan', 'education', 'career']) {
    if (!p[f]) throw new Error(`Missing field: ${f}`);
  }
  return p;
}

// ─── USER REGISTRATION ────────────────────────────────────────────────────────
async function upsertUser(user_id, username, full_name) {
  const { error } = await supabase.from('users').upsert({
    user_id: String(user_id),
    username: username || null,
    full_name: full_name || null,
    updated_at: new Date().toISOString(),
  }, { onConflict: 'user_id', ignoreDuplicates: false });
  if (error) console.warn('User upsert warning:', error.message);
}

// ─── COMPANY SECTION ──────────────────────────────────────────────────────────
function buildCompanySection(profile) {
  const out = [];
  const cn = profile.company_name;

  out.push(centerPara([
    run('ИНФОРМАЦИЯ', { bold: true, color: RED }),
    new TextRun({ break: 1 }),
    run('о компании ', { bold: true }),
    run('«', { bold: true }),
    run(cn, { bold: true, color: BLUE }),
    run('»', { bold: true }),
  ], { spacing: SP_COMPACT }));

  // Company logo centered
  if (profile.logoImg) {
    const logoPara = makeImagePara(profile.logoImg, 5, 2);
    if (logoPara) out.push(logoPara);
  }
  out.push(ep());

  // Company description
  const sentences = profile.company_description.split(/(?<=[.!?])\s+/);
  const chunks = [];
  let chunk = '';
  for (const s of sentences) {
    chunk += (chunk ? ' ' : '') + s;
    if (chunk.length > 250) { chunks.push(chunk); chunk = ''; }
  }
  if (chunk) chunks.push(chunk);
  for (const c of chunks) {
    out.push(para(descRuns(c, cn), { indent: true }));
  }

  out.push(ep());

  // Uzbekistan section
  out.push(centerPara([
    run('Информация о деятельности компании', { bold: true }),
    new TextRun({ break: 1 }),
    run('«', { bold: true }),
    run(cn, { bold: true, color: BLUE }),
    run('» в Республике Узбекистан', { bold: true }),
  ], { spacing: SP_COMPACT }));

  out.push(ep());

  if (profile.company_intro) {
    out.push(para(descRuns(profile.company_intro, cn), { indent: true }));
  }

  for (const s of profile.company_uzbekistan) {
    out.push(para([
      run(s.sector + ': ', { bold: true }),
      ...descRuns(s.description, cn),
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

function uzbekCell(nameRaw, width) {
  if (!nameRaw) return tc([ep()], width);
  const parts = nameRaw.split('\n');
  const words = (parts[0] || '').trim().split(/\s+/);
  const surname = words[0] || '';
  const firstname = words.slice(1).join(' ');
  const titlePart = parts.slice(1).join(' ') || '';

  const paras = [
    new Paragraph({
      children: [
        run(surname.toUpperCase(), { bold: true, color: BLUE }),
        ...(firstname ? [new TextRun({ break: 1 }), run(firstname, { color: BLUE })] : []),
      ],
      spacing: SP_COMPACT,
    }),
    epTiny(),
    ...(titlePart ? [new Paragraph({ children: [run(titlePart)], spacing: SP_COMPACT, alignment: AlignmentType.BOTH })] : []),
  ];

  return tc(paras, width);
}

function foreignCell(profile, width) {
  if (!profile) return tc([ep()], width);
  return tc([
    new Paragraph({
      children: [
        run(profile.surname_caps, { bold: true, color: BLUE }),
        new TextRun({ break: 1 }),
        run(profile.first_name, { color: BLUE }),
      ],
      spacing: SP_COMPACT,
    }),
    epTiny(),
    new Paragraph({ children: [run(profile.title)], spacing: SP_COMPACT, alignment: AlignmentType.BOTH }),
  ], width);
}

function buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles) {
  const out = [];
  out.push(centerPara([run('СПИСОК УЧАСТНИКОВ', { bold: true })], { spacing: SP_COMPACT }));
  out.push(ep());

  const dateStr = formatDate(meetingDate, meetingTime);
  const numW = 400;
  const colW = Math.floor((CW - numW) / 2);

  const rows = [
    new TableRow({ children: [
      tc([para([run('Дата: ', { bold: true }), ...yearRuns(dateStr)])], CW, { span: 3 }),
    ]}),
    new TableRow({ children: [tc([ep()], numW), tc([ep()], colW), tc([ep()], colW)] }),
    new TableRow({ children: [
      tc([ep()], numW),
      tc([centerPara([run('От узбекской стороны', { bold: true })], { spacing: SP_COMPACT })], colW),
      tc([centerPara([run('От иностранной стороны', { bold: true })], { spacing: SP_COMPACT })], colW),
    ]}),
    ...[0,1,2,3].map(i => new TableRow({ children: [
      tc([centerPara([run(`${i+1}.`, { bold: true })], { spacing: SP_COMPACT })], numW),
      uzbekCell(uzbekParticipants[i] || '', colW),
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

  out.push(centerPara([run(profile.full_name_bio, { bold: true, color: BLUE })], { spacing: SP_COMPACT }));

  // Person photo centered (passport size ~4cm wide, 5cm tall)
  if (profile.personImg) {
    const photoPara = makeImagePara(profile.personImg, 4, 5);
    if (photoPara) out.push(photoPara);
  }

  out.push(centerPara([
    run(profile.title, { bold: true }),
    new TextRun({ break: 1 }),
    run('«'),
    run(profile.company_name, { bold: true, color: BLUE }),
    run('»'),
  ], { spacing: SP_COMPACT }));

  out.push(ep());

  const col1 = 2269;
  const col2 = 426;
  const col3 = CW - col1 - col2;

  // Education table
  const eduRows = [
    new TableRow({ children: [
      tc([para([run('Образование:', { bold: true })], { spacing: SP_BODY })], CW, { span: 3 }),
    ]}),
    ...profile.education.map(e => {
      const descChildren = [
        ...(e.degree ? [run(e.degree + ' ')] : []),
        ...(e.institution ? [run('«'), run(e.institution, { color: BLUE }), run('»')] : []),
        ...(e.country ? [run(' (', { italic: true }), run(e.country, { italic: true, color: BLUE }), run(')', { italic: true })] : []),
      ];
      return new TableRow({ children: [
        tc([para(yearRuns(e.years), { spacing: SP_BODY })], col1),
        tc([para([run('-', { italic: true })], { spacing: SP_BODY, align: AlignmentType.CENTER })], col2),
        tc([para(descChildren.length ? descChildren : [run(e.degree || '')], { spacing: SP_BODY })], col3),
      ]});
    }),
  ];

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: eduRows }));
  out.push(ep());

  // Career table
  const careerRows = [
    new TableRow({ children: [
      tc([para([run('Профессиональная деятельность:', { bold: true })], { spacing: SP_BODY })], CW, { span: 3 }),
    ]}),
    ...profile.career.map(c => {
      const descChildren = [
        ...(c.role ? [run(c.role + ' ')] : []),
        ...(c.company ? [run('«'), run(c.company, { color: BLUE }), run('»')] : []),
        ...(c.country ? [run(' (', { italic: true }), run(c.country, { italic: true, color: BLUE }), run(')', { italic: true })] : []),
      ];
      return new TableRow({ children: [
        tc([para(yearRuns(c.years), { spacing: SP_BODY })], col1),
        tc([para([run('-', { italic: true })], { spacing: SP_BODY, align: AlignmentType.CENTER })], col2),
        tc([para(descChildren.length ? descChildren : [run(c.role || '')], { spacing: SP_BODY })], col3),
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

// ─── ENDPOINTS ────────────────────────────────────────────────────────────────

// Generate document
app.post('/generate', async (req, res) => {
  const { inputs, meeting_date, meeting_time, uzbek_participants, user_id, username, full_name } = req.body;

  if (!inputs?.length) return res.status(400).json({ error: 'inputs required' });
  if (!meeting_date) return res.status(400).json({ error: 'meeting_date required' });
  if (!uzbek_participants?.length) return res.status(400).json({ error: 'uzbek_participants required' });
  if (!user_id) return res.status(400).json({ error: 'user_id required' });

  try {
    // Register/update user
    await upsertUser(user_id, username, full_name);

    console.log(`Researching ${inputs.length} person(s) for user ${user_id}...`);
    const rawProfiles = await Promise.all(inputs.map(i => researchPerson(i.query)));
    const profiles = rawProfiles.map(validateProfile);
    console.log('Profiles validated. Generating DOCX...');

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
      doc_id,
      user_id: String(user_id),
      download_url,
      meeting_date,
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

// Get user profile + document history
app.get('/profile/:user_id', async (req, res) => {
  const { user_id } = req.params;
  try {
    const { data: user, error: userErr } = await supabase
      .from('users')
      .select('*')
      .eq('user_id', user_id)
      .single();
    if (userErr) return res.status(404).json({ error: 'User not found' });

    const { data: docs, error: docsErr } = await supabase
      .from('documents')
      .select('doc_id, download_url, meeting_date, profiles_count, queries, created_at')
      .eq('user_id', user_id)
      .order('created_at', { ascending: false })
      .limit(20);

    return res.json({ user, documents: docs || [] });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// Admin: get all documents
app.get('/admin/documents', async (req, res) => {
  if (req.headers['x-admin-key'] !== ADMIN_KEY) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  try {
    const { data, error } = await supabase
      .from('documents')
      .select('*, users(full_name, username)')
      .order('created_at', { ascending: false })
      .limit(100);
    if (error) throw error;
    return res.json({ documents: data });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// Admin: get all users
app.get('/admin/users', async (req, res) => {
  if (req.headers['x-admin-key'] !== ADMIN_KEY) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  try {
    const { data, error } = await supabase
      .from('users')
      .select('*')
      .order('created_at', { ascending: false });
    if (error) throw error;
    return res.json({ users: data });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

app.get('/health', (_, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`KimKim API running on port ${PORT}`));
