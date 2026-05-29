const {
  Document, Packer, Paragraph, TextRun, TabStopType, TabStopPosition,
  AlignmentType, LevelFormat, BorderStyle, UnderlineType
} = require('docx');
const fs = require('fs');

const NAVY = "1B3A6B";
const BLUE = "2E74B5";
const BLACK = "000000";
const DARK = "1A1A1A";

const PAGE_W = 12240;
const PAGE_H = 15840;
const MARGIN = 900; // 0.625 inch
const CONTENT_W = PAGE_W - (MARGIN * 2);

// ── helpers ──────────────────────────────────────────────────────────────────

const gap = (pts = 80) => new Paragraph({ children: [new TextRun("")], spacing: { before: 0, after: pts } });

const rule = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
  spacing: { before: 0, after: 80 },
  children: [new TextRun("")]
});

const sectionHeader = (text) => new Paragraph({
  spacing: { before: 160, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 3 } },
  children: [new TextRun({ text, bold: true, color: BLUE, size: 22, font: "Arial", allCaps: true })]
});

const jobHeader = (title, dates) => new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
  spacing: { before: 120, after: 40 },
  children: [
    new TextRun({ text: title, bold: true, size: 20, font: "Arial", color: DARK }),
    new TextRun({ text: "\t" + dates, size: 20, font: "Arial", color: DARK })
  ]
});

const bodyText = (text) => new Paragraph({
  spacing: { before: 0, after: 60 },
  children: [new TextRun({ text, size: 19, font: "Arial", color: DARK })]
});

const bullet = (parts) => {
  // parts: array of { text, bold? }
  const runs = parts.map(p => new TextRun({
    text: p.text,
    bold: p.bold || false,
    size: 19,
    font: "Arial",
    color: DARK
  }));
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 0, after: 60 },
    children: runs
  });
};

const starBullet = (parts) => {
  const runs = [
    new TextRun({ text: "★ ", bold: true, size: 19, font: "Arial", color: DARK }),
    ...parts.map(p => new TextRun({ text: p.text, bold: p.bold || false, size: 19, font: "Arial", color: DARK }))
  ];
  return new Paragraph({
    indent: { left: 360, hanging: 360 },
    spacing: { before: 0, after: 60 },
    children: runs
  });
};

// ── document ─────────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 300 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    children: [

      // ── NAME ──
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          new TextRun({ text: "JESSICA CARLIN", bold: true, size: 48, font: "Arial", color: NAVY }),
          new TextRun({
            text: "\tLondon • linkedin.com/in/jessicacarlin",
            size: 18, font: "Arial", color: DARK
          })
        ]
      }),
      new Paragraph({
        spacing: { before: 0, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          new TextRun({ text: "", size: 48 }),
          new TextRun({ text: "\t07792387454 • hello@jesscarlin.com", size: 18, font: "Arial", color: DARK })
        ]
      }),

      rule(),

      // ── TAGLINE ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 80 },
        children: [new TextRun({
          text: "Strategic Partnerships · Marketing Innovation · Cross-Sector Transformation",
          bold: true, italics: true, size: 20, font: "Arial", color: DARK
        })]
      }),

      // ── SUMMARY ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({
          text: "Senior partnerships / marketing executive specialising in the intersection of creative content, technology, and commerce. Proven track record in venture creation, operational transformation, building strategic relationships that drive innovation across fragmented creative ecosystems, and delivering commercial growth through AI and emerging technology.",
          italics: true, size: 19, font: "Arial", color: DARK
        })]
      }),

      // ── KEY ACCOMPLISHMENTS ──
      sectionHeader("KEY ACCOMPLISHMENTS"),
      gap(60),

      starBullet([
        { text: "Venture creation & growth", bold: true },
        { text: " — Co-founded AdTech/data startup; secured 2x Innovate UK grants; delivered measurable box-office uplifts for studio titles; launched new products for independent cinemas/SMEs." }
      ]),
      starBullet([
        { text: "Innovation deployment", bold: true },
        { text: " — Led AI/automation roadmaps and martech integrations for broadcasters/streamers and media orgs; shipped case studies with major partners including Peacock, YES Network, Disney, Warner Bros, and more." }
      ]),
      starBullet([
        { text: "Industry Education & Upskilling", bold: true },
        { text: " — Created Early Adoptr podcast + workshops to upskill SMEs/creatives on AI and responsible adoption of emerging technologies." }
      ]),

      gap(60),

      // ── SELECT PROFESSIONAL EXPERIENCE ──
      sectionHeader("SELECT PROFESSIONAL EXPERIENCE"),

      // Early Adoptr
      jobHeader("Creator & Co-Host - Early Adoptr", "May 2025 - present"),
      bodyText("Created and co-host platform simplifying AI and emerging technology adoption for SMEs and creative organisations, available across all major podcast platforms, including YouTube."),
      bullet([
        { text: "Produced 15+ episodes reaching thousands of listeners,", bold: true },
        { text: " covering topics from LLMs to prompt engineering to AI in the creative industries to synthetic data, all aimed at building practical knowledge to address the gap between complex emerging technologies and SMEs." }
      ]),
      bullet([{ text: "Built strategic industry partnerships through podcast platform, securing AI workshop delivery for Cineworld Cinemas, and panels for AI in cinemas at the UK Cinema Association conference, demonstrating ability to scale market opportunities and bridge technology with creative applications." }]),
      bullet([{ text: "Established expertise in responsible technology adoption featuring industry leaders including Sean Bhardwaj of Breakthrough Growth Partners for specialised AI policy series, addressing digital ethics and responsible innovation implementation." }]),

      // Atonik
      jobHeader("Partner - Atonik", "March 2025 - present"),
      bodyText("Strategic advisory role within media and entertainment consultancy, providing expertise on the intersection of content, commerce, and technology for industry clients and business development opportunities, with an emphasis on marketing and technology."),
      bullet([{ text: "Developed “So This Happened...” newsletter analysing news stories at the intersection of content, commerce and technology, identifying trend development and strategic implications for entertainment, media, and tech sectors—supporting innovation capacity across creative ecosystems." }]),
      bullet([{ text: "Advised on strategic projects including new media & entertainment conference development, comprehensive RFI for a major UK broadcaster measurement initiative, and global growth strategy for global news platform, working across traditionally fragmented creative subsectors." }]),

      // Jessica Carlin Consulting
      jobHeader("Principal Consultant - Jessica Carlin Consulting", "March 2025 - present"),
      bodyText("Independent consultancy providing digital transformation advisory services for creative and media organisations, focusing on operational resilience, cross-functional integration, and emerging technology adoption."),

      // ── EITCC BULLETS (new) ──
      bullet([
        { text: "Led a two-phase marketing transformation engagement for EIT Culture & Creativity", bold: true },
        { text: " (major EU-funded body, 50+ staff): Phase 1 comprised a comprehensive 6-week capability audit across communications, technology, and operational processes, delivering a prioritised 30/60/90-day roadmap and KPI framework (£10,800)." }
      ]),
      bullet([
        { text: "Currently leading Phase 2 implementation", bold: true },
        { text: " across three parallel tracks: operational quick wins (Monday.com workflow design, centralised materials hub, communications calendar); foundation building (AI policy development, AI knowledge hub, style guides, approval workflow redesign); and technology stack evaluation and implementation (CRM selection, Slack, Zapier/Make automation)." }
      ]),

      // existing JCC bullets
      bullet([
        { text: "Led comprehensive transformation audit for The Economist Group", bold: true },
        { text: " across growth marketing and creative operations, interviewing 20+ team members across 5 business units to assess current systems and identify optimisation opportunities for sustainable growth." }
      ]),
      bullet([{ text: "Delivered cross-functional technology integration through successful Airtable implementations and data migration across creative teams, addressing skills gaps and innovation capacity challenges faced by creative organisations adapting to digital-first operations." }]),
      bullet([{ text: "Developed strategic AI implementation roadmap focusing on operational efficiency, cross-functional collaboration, and future-proofing creative workflows—supporting organisations navigating the intersection of creativity and emerging technology." }]),

      // Pickaxe
      jobHeader("Director, Marketing & Partnerships – Pickaxe Foundry", "2021 – Feb 2025"),
      bodyText("Led all aspects of marketing & partnership strategy for the digital transformation firm. Managed large-scale projects to improve internal and client marketing operations and build partnership opportunities, supporting innovation implementation across the media sector."),
      bullet([{ text: "Delivered transformative partnership strategy managing relationships with 10+ tech vendors including securing Airtable Partner Program launch, featured Peacock case study with Braze & mParticle, and ITV-Amplitude campaign, demonstrating cross-sector collaboration capabilities within the creative industry." }]),
      bullet([{ text: "Led operational efficiency transformation with Peacock, implementing a tracking platform review and delivering improved marketing operations efficiency praised by client leadership—addressing innovation implementation challenges." }]),
      bullet([{ text: "Developed and implemented a new marketing strategy delivering outstanding growth in marketing effectiveness, resulting in speaking appearances at IBC and the Peacock MarTech Summit, 2x finalist nominations for the Digiday Technology Awards, silver at The Drum Awards, and a 50% increase in content creation YOY." }]),

      // Screen Moguls
      jobHeader("Co-founder / CMO / CCO – Screen Moguls", "2018 – 2021"),
      bodyText("Led the data and AdTech startup, including developing and overseeing the commercial and marketing functions. Built and coached a team of marketing, creative, and development professionals. Created brand marketing and partnership vision. Oversaw data initiatives and marketing campaigns for film studios, media agencies, and cinema chains."),
      bullet([{ text: "Enhanced Screen Moguls’ presence and marketing technology by building strategic partnerships, including a pioneering partnership to make Screen Moguls one of the first companies globally to integrate TikTok’s API." }]),
      bullet([{ text: "Developed and ran innovative data campaigns with Universal that delivered a ~16% audience uplift and ~£37k incremental revenue across 4 titles, with a 29% audience uplift and £117k incremental box office revenue for Downton Abbey." }]),
      bullet([{ text: "Transformed the business by designing an innovative AdTech solution that both simplified multi-channel social media marketing for independent cinemas and SMEs." }]),

      // RealD
      jobHeader("Marketing Manager, EMEA – RealD", "2012 – 2017"),
      bodyText("Oversaw marketing strategy, delivery, and execution for all B2B and B2C initiatives across 20+ EMEA territories, with a goal to increase revenues by driving consumer demand and ticket sales for premium 3D cinema. Led a globally dispersed marketing, partnerships and PR team. Executed promotional activities and content partnerships for 40+ 3D movies per year, including iconic franchises such as Marvel and Star Wars."),

      gap(80),

      // ── EDUCATION ──
      sectionHeader("EDUCATION AND PROFESSIONAL DEVELOPMENT"),
      gap(60),

      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Master of Arts, History – University of St Andrews", bold: true, size: 19, font: "Arial", color: DARK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "2x Innovate UK Grant Recipient", bold: true, size: 19, font: "Arial", color: DARK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Graduate - London & Partners Business Growth Program", bold: true, size: 19, font: "Arial", color: DARK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Certified in HTML / CSS / Javascript / React / Python", bold: true, size: 19, font: "Arial", color: DARK })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/sessions/beautiful-dazzling-noether/mnt/outputs/jessica-carlin-cv.docx', buffer);
  console.log('Done');
});
