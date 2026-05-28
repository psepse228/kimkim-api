const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
        PageBreak, HeadingLevel } = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
app.use(express.json());

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
  const searchText = await tavilySearch(`${query} biography career education`);
  const companyQuery = query.includes('http') ? query : `${query} company`;
  const companyText = await tavilySearch(`${companyQuery} Uzbekistan cooperation investment`);

  const prompt = `You are extracting structured data about a foreign business person for an Uzbek Ministry meeting document.

Research data about the person:
${searchText}

Research data about their company and Uzbekistan:
${companyText}

Return ONLY a JSON object with this exact structure. All text fields must be in Russian. If information is not found, use "Информация не найдена". Never return null or empty strings.

{
  "full_name": "ФАМИЛИЯ Имя (surname in caps, then first name)",
  "title": "Full job title at company in Russian",
  "company_name": "Company name in original language",
  "company_description": "3-4 sentences about the company: what it does, size, headquarters, history. In Russian.",
  "company_uzbekistan": [
    { "sector": "Sector name in Russian", "description": "What the company does in this sector in Uzbekistan. 2-3 sentences." }
  ],
  "education": [
    { "years": "YYYY-YYYY гг.", "description": "Degree and institution name. In Russian." }
  ],
  "career": [
    { "years": "YYYY-YYYY гг. or с YYYY года", "description": "Role and company name. In Russian." }
  ]
}

Requirements:
- full_name: surname ALL CAPS, then first name normal case
- company_uzbekistan: minimum 2 sectors, maximum 5
- education: chronological order, minimum 1 entry
- career: chronological order, minimum 3 entries
- All years in Russian format: "1999-2007 гг." or "с 2021 года"`;

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${OPENAI_API_KEY}`
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      max_tokens: 2000,
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

// ─── DOCX GENERATION ─────────────────────────────────────────────────────────

const FONT = 'Times New Roman';
const PAGE_WIDTH = 11906; // A4
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 9026

const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function txt(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: opts.size || 24,
    bold: opts.bold || false,
    italics: opts.italic || false,
    ...opts
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { after: opts.after !== undefined ? opts.after : 120, before: opts.before || 0 },
    alignment: opts.align || AlignmentType.BOTH,
    ...opts
  });
}

function emptyPara() {
  return para([txt('')], { after: 80 });
}

function pageBreakPara() {
  return new Paragraph({ children: [new PageBreak()] });
}

function cell(content, opts = {}) {
  const children = Array.isArray(content) ? content : [
    para(Array.isArray(content) ? content : [typeof content === 'string' ? txt(content, opts.textOpts || {}) : content], {
      after: 60,
      align: opts.align || AlignmentType.LEFT
    })
  ];
  return new TableCell({
    borders: opts.noBorder ? noBorders : borders,
    width: { size: opts.width || Math.floor(CONTENT_WIDTH / 2), type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span || 1,
    children: Array.isArray(content) && content[0] instanceof Paragraph ? content : [
      para([typeof content === 'string' ? txt(content, opts.textOpts || {}) : content], {
        after: 60,
        align: opts.align || AlignmentType.LEFT
      })
    ]
  });
}

function buildCompanySection(profile) {
  const children = [];

  // Company info title
  children.push(para([
    txt('ИНФОРМАЦИЯ ', { bold: true }),
    txt('о деятельности ', { bold: true }),
    txt(`«${profile.company_name}»`, { bold: true })
  ], { after: 160, align: AlignmentType.CENTER }));

  // Company description
  children.push(para([txt(profile.company_description)], { after: 160 }));

  // Uzbekistan activity title
  children.push(para([
    txt('ДЕЯТЕЛЬНОСТЬ', { bold: true }),
    txt(` «${profile.company_name}» `, { bold: true }),
    txt('в Узбекистане', { bold: true })
  ], { after: 160, align: AlignmentType.CENTER }));

  // Sector paragraphs
  for (const sector of profile.company_uzbekistan) {
    children.push(para([
      txt(`${sector.sector}: `, { bold: true }),
      txt(sector.description)
    ], { after: 120 }));
  }

  return children;
}

function buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles) {
  const children = [];

  children.push(para([txt('СПИСОК УЧАСТНИКОВ', { bold: true })],
    { after: 120, align: AlignmentType.CENTER }));

  // Format date
  const months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
                  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const d = new Date(meetingDate);
  const dateStr = `${d.getDate()} ${months[d.getMonth()]} т.г., ${meetingTime}`;

  const colW = Math.floor(CONTENT_WIDTH / 2);

  const rows = [];

  // Row 1: date merged
  rows.push(new TableRow({
    children: [
      new TableCell({
        borders,
        columnSpan: 3,
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [para([txt(`Дата: ${dateStr}`, { bold: true })], { after: 60 })]
      })
    ]
  }));

  // Row 2: empty spacer
  rows.push(new TableRow({
    children: [
      new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [emptyPara()] }),
      new TableCell({ borders, width: { size: colW - 200, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [emptyPara()] }),
      new TableCell({ borders, width: { size: colW - 200, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [emptyPara()] })
    ]
  }));

  // Row 3: headers
  rows.push(new TableRow({
    children: [
      new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 }, children: [emptyPara()] }),
      new TableCell({ borders, width: { size: colW - 200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt('От узбекской стороны', { bold: true })], { after: 60, align: AlignmentType.CENTER })] }),
      new TableCell({ borders, width: { size: colW - 200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt('От иностранной стороны', { bold: true })], { after: 60, align: AlignmentType.CENTER })] })
    ]
  }));

  // Data rows (max 4)
  const maxRows = 4;
  for (let i = 0; i < maxRows; i++) {
    const uzbek = uzbekParticipants[i] || '';
    const foreignProfile = profiles[i] || null;

    const foreignContent = foreignProfile
      ? [
          txt(`${foreignProfile.full_name}`, { bold: true }),
          txt('\n'),
          txt(foreignProfile.title)
        ]
      : [txt('')];

    rows.push(new TableRow({
      children: [
        new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 }, children: [para([txt(`${i + 1}.`, { bold: true })], { after: 60, align: AlignmentType.CENTER })] }),
        new TableCell({ borders, width: { size: colW - 200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(uzbek, { bold: !!uzbek })], { after: 60 })] }),
        new TableCell({ borders, width: { size: colW - 200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para(foreignContent, { after: 60 })] })
      ]
    }));
  }

  children.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [400, colW - 200, colW - 200],
    rows
  }));

  return children;
}

function buildBioSection(profile) {
  const children = [];

  children.push(pageBreakPara());

  // Person title
  children.push(para([txt(profile.full_name, { bold: true })],
    { after: 40, align: AlignmentType.CENTER }));
  children.push(para([txt(profile.title, { bold: true })],
    { after: 160, align: AlignmentType.CENTER }));

  const col1 = 2000;
  const col2 = 400;
  const col3 = CONTENT_WIDTH - col1 - col2;

  // Education table
  const eduRows = [
    new TableRow({
      children: [
        new TableCell({
          borders,
          columnSpan: 3,
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          shading: { fill: 'D9D9D9', type: ShadingType.CLEAR },
          children: [para([txt('Образование:', { bold: true })], { after: 60 })]
        })
      ]
    }),
    ...profile.education.map(e => new TableRow({
      children: [
        new TableCell({ borders, width: { size: col1, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(e.years)], { after: 60 })] }),
        new TableCell({ borders, width: { size: col2, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 }, children: [para([txt('-', { italic: true })], { after: 60, align: AlignmentType.CENTER })] }),
        new TableCell({ borders, width: { size: col3, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(e.description)], { after: 60 })] })
      ]
    }))
  ];

  children.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col1, col2, col3],
    rows: eduRows
  }));

  children.push(emptyPara());

  // Career table
  const careerRows = [
    new TableRow({
      children: [
        new TableCell({
          borders,
          columnSpan: 3,
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          shading: { fill: 'D9D9D9', type: ShadingType.CLEAR },
          children: [para([txt('Профессиональная деятельность:', { bold: true })], { after: 60 })]
        })
      ]
    }),
    ...profile.career.map(c => new TableRow({
      children: [
        new TableCell({ borders, width: { size: col1, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(c.years)], { after: 60 })] }),
        new TableCell({ borders, width: { size: col2, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 }, children: [para([txt('-', { italic: true })], { after: 60, align: AlignmentType.CENTER })] }),
        new TableCell({ borders, width: { size: col3, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(c.description)], { after: 60 })] })
      ]
    }))
  ];

  children.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col1, col2, col3],
    rows: careerRows
  }));

  return children;
}

async function generateDocx(profiles, meetingDate, meetingTime, uzbekParticipants) {
  let allChildren = [];

  // Company sections for all profiles
  for (let i = 0; i < profiles.length; i++) {
    const companyChildren = buildCompanySection(profiles[i]);
    allChildren = allChildren.concat(companyChildren);
    if (i < profiles.length - 1) allChildren.push(pageBreakPara());
  }

  allChildren.push(pageBreakPara());

  // Participants table
  const tableChildren = buildParticipantsTable(meetingDate, meetingTime, uzbekParticipants, profiles);
  allChildren = allChildren.concat(tableChildren);

  // Bio sections
  for (const profile of profiles) {
    const bioChildren = buildBioSection(profile);
    allChildren = allChildren.concat(bioChildren);
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
  if (!uzbek_participants || !uzbek_participants.length) return res.status(400).json({ error: 'uzbek_participants required' });

  try {
    // 1. Research all persons in parallel
    console.log(`Researching ${inputs.length} person(s)...`);
    const rawProfiles = await Promise.all(inputs.map(i => researchPerson(i.query)));

    // 2. Validate
    const profiles = rawProfiles.map(validateProfile);
    console.log('All profiles validated.');

    // 3. Generate DOCX
    console.log('Generating DOCX...');
    const docBuffer = await generateDocx(profiles, meeting_date, meeting_time || '11:00', uzbek_participants);

    // 4. Upload to Supabase Storage
    const doc_id = `doc_${Date.now()}_${user_id}`;
    const fileName = `${doc_id}.docx`;
    console.log(`Uploading ${fileName} to Supabase...`);

    const { error: uploadError } = await supabase.storage
      .from('documents')
      .upload(fileName, docBuffer, {
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        upsert: true
      });

    if (uploadError) throw new Error(`Supabase upload failed: ${uploadError.message}`);

    const download_url = `${SUPABASE_URL}/storage/v1/object/public/documents/${fileName}`;

    // 5. Save to DB
    console.log('Saving to Supabase DB...');
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
    console.error('Generation error:', err.message);
    return res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/health', (_, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`KimKim API running on port ${PORT}`));
