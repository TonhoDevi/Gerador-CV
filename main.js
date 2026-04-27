const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  LevelFormat, ExternalHyperlink, HeadingLevel
} = require('docx');
const fs = require('fs');

// ─── Colors ────────────────────────────────────────────────
const ACCENT   = "1A56DB";   // deep blue
const DARK     = "1E293B";   // near-black text
const MID      = "475569";   // secondary text
const LIGHT    = "EFF6FF";   // sidebar bg
const WHITE    = "FFFFFF";
const RULE     = "BFDBFE";   // subtle rule color

// ─── Sizes (half-points) ───────────────────────────────────
const S_NAME   = 28;  // 20pt
const S_TITLE  = 22;  // 11pt
const S_SECTION= 22;  // 11pt – section headers
const S_BODY   = 18;  // 9pt
const S_SMALL  = 16;  // 8pt

// ─── Page / column metrics (DXA) ──────────────────────────
const PAGE_W   = 11906;
const PAGE_H   = 16838;
const MAR_T    = 720;
const MAR_B    = 720;
const MAR_L    = 0;
const MAR_R    = 0;
const COL_LEFT = 2900;   // sidebar
const COL_RIGHT= 9006;   // main   (11906 - 900 gutter = 9006, but we use padding)

// ─── Helper: no border ────────────────────────────────────
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ─── Helper: rule paragraph ───────────────────────────────
function rule(color = RULE) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color, space: 1 } },
    children: []
  });
}

// ─── Helper: section header (main column) ─────────────────
function sectionHeader(text) {
  return [
    new Paragraph({
      spacing: { before: 220, after: 40 },
      children: [
        new TextRun({ text: text.toUpperCase(), bold: true, size: S_SECTION, color: ACCENT, font: "Arial" })
      ]
    }),
    rule(RULE)
  ];
}

// ─── Helper: sidebar section header ───────────────────────
function sideHeader(text) {
  return [
    new Paragraph({
      spacing: { before: 220, after: 40 },
      children: [
        new TextRun({ text: text.toUpperCase(), bold: true, size: S_SMALL, color: WHITE, font: "Arial" })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: "93C5FD", space: 1 } },
      children: []
    })
  ];
}

// ─── Helper: sidebar body text ────────────────────────────
function sideText(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text, size: S_SMALL, color: WHITE, font: "Arial", ...opts })
    ]
  });
}

// ─── Helper: body paragraph ───────────────────────────────
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text, size: S_BODY, color: DARK, font: "Arial", ...opts })
    ]
  });
}

// ─── Helper: bullet item ──────────────────────────────────
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text, size: S_BODY, color: DARK, font: "Arial" })
    ]
  });
}

// ─── Helper: experience block ─────────────────────────────
function expBlock(role, company, period, location, bullets) {
  return [
    new Paragraph({
      spacing: { before: 140, after: 20 },
      children: [
        new TextRun({ text: role, bold: true, size: S_BODY + 2, color: DARK, font: "Arial" }),
        new TextRun({ text: "   ·   ", size: S_BODY, color: MID, font: "Arial" }),
        new TextRun({ text: company, size: S_BODY, color: ACCENT, font: "Arial", bold: true })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: `${period}  |  ${location}`, size: S_SMALL, color: MID, font: "Arial", italics: true })
      ]
    }),
    ...bullets.map(b => bullet(b))
  ];
}

// ─── Helper: project block ────────────────────────────────
function projectBlock(name, url, desc, stack) {
  const titleChildren = [
    new TextRun({ text: name, bold: true, size: S_BODY + 2, color: DARK, font: "Arial" })
  ];
  if (url) {
    titleChildren.push(new TextRun({ text: "   ", size: S_BODY, font: "Arial" }));
    titleChildren.push(new ExternalHyperlink({
      link: url,
      children: [new TextRun({ text: url, size: S_SMALL, color: ACCENT, font: "Arial", style: "Hyperlink" })]
    }));
  }
  return [
    new Paragraph({ spacing: { before: 140, after: 20 }, children: titleChildren }),
    new Paragraph({
      spacing: { before: 0, after: 30 },
      children: [new TextRun({ text: desc, size: S_BODY, color: DARK, font: "Arial" })]
    }),
    new Paragraph({
      spacing: { before: 20, after: 80 },
      children: [
        new TextRun({ text: "Stack: ", bold: true, size: S_SMALL, color: MID, font: "Arial" }),
        new TextRun({ text: stack, size: S_SMALL, color: MID, font: "Arial" })
      ]
    })
  ];
}

// ─── SIDEBAR content ──────────────────────────────────────
const sidebarChildren = [
  // Name block (white bg header area baked into sidebar)
  new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [
      new TextRun({ text: "ANTONIO HENRIQUE", bold: true, size: S_NAME, color: WHITE, font: "Arial" })
    ]
  }),
  new Paragraph({
    spacing: { before: 0, after: 20 },
    children: [
      new TextRun({ text: "BATISTA DO NASCIMENTO", bold: true, size: S_NAME - 6, color: "BFDBFE", font: "Arial" })
    ]
  }),
  new Paragraph({
    spacing: { before: 0, after: 300 },
    children: [
      new TextRun({ text: "Desenvolvedor Full Stack", size: S_TITLE, color: "93C5FD", font: "Arial", italics: true })
    ]
  }),

  // Contact
  ...sideHeader("Contato"),
  sideText("📍 Recife, PE — Brasil"),
  sideText("📞 (81) 98422-4699"),
  sideText("✉ antoniohenriquebn@hotmail.com"),
  sideText("🔗 github.com/TonhoDevi"),
  sideText("🔗 linkedin.com/in/antonio-henrique-bn"),

  // Skills
  ...sideHeader("Tecnologias"),
  sideText("Java  ·  Spring Boot  ·  JPA/Hibernate"),
  sideText("Python  ·  JavaScript  ·  HTML/CSS"),
  sideText("C#  ·  SQL  ·  PostgreSQL"),
  sideText("REST APIs  ·  Git / GitHub"),
  sideText("Supabase  ·  Vercel  ·  Claude API"),

  // Languages
  ...sideHeader("Idiomas"),
  sideText("Português", { bold: true }),
  sideText("Nativo"),
  new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }),
  sideText("Inglês", { bold: true }),
  sideText("Avançado — leitura, escrita e"),
  sideText("compreensão oral completas;"),
  sideText("conversação intermediária"),
  new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }),
  sideText("Espanhol", { bold: true }),
  sideText("Básico"),

  // Soft Skills
  ...sideHeader("Competências"),
  sideText("• Resolução de problemas"),
  sideText("• Criatividade"),
  sideText("• Didática e comunicação"),
  sideText("• Trabalho em equipe"),
  sideText("• Autonomia e proatividade"),

  // Formation (sidebar)
  ...sideHeader("Formação"),
  sideText("Engenharia de Software", { bold: true }),
  sideText("Estácio · 2025 – Cursando"),
  new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }),
  sideText("Desenv. Java Full Stack", { bold: true }),
  sideText("Fuctura Tecnologia"),
  sideText("Jan 2025 – Nov 2025"),
  new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }),
  sideText("Gestão Financeira", { bold: true }),
  sideText("Unifateci · Ago 2023 – Ago 2025"),
];

// ─── MAIN COLUMN content ──────────────────────────────────
const mainChildren = [
  // Profile
  ...sectionHeader("Perfil Profissional"),
  new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [
      new TextRun({
        text: "Desenvolvedor Full Stack Júnior com sólida experiência prática em Java/Spring Boot, JavaScript, HTML/CSS e PostgreSQL. Atuou como professor de programação na Fuctura Tecnologia, conduzindo 6 turmas e formando mais de 90 alunos em Java e Python — o que demonstra domínio técnico aliado à capacidade de comunicar e ensinar código com clareza. Participou como desenvolvedor em projeto real de produção (sistema de agendamento do BOPE), com backend em Java/Spring Boot, integração com PostgreSQL e consumo de APIs no frontend. Possui projetos próprios em produção com stacks modernas (Supabase, Vercel, Claude API). Inglês avançado para documentação técnica, escrita e comunicação oral.",
        size: S_BODY, color: DARK, font: "Arial"
      })
    ]
  }),

  // Experience
  ...sectionHeader("Experiência"),
  ...expBlock(
    "Professor de Programação & Desenvolvedor Júnior",
    "Fuctura Tecnologia",
    "Set 2025 – Presente",
    "Recife, PE",
    [
      "Conduzi 6 turmas de Java e Python (90+ alunos), criando conteúdo didático, exercícios práticos e acompanhando o desenvolvimento individual de cada aluno.",
      "Contribuí no desenvolvimento do projeto oficial do BOPE: backend Java/Spring Boot, integração com PostgreSQL, consumo de APIs no frontend e construção do fluxo principal de agendamento de turnos.",
      "Integro a API da Anthropic (Claude) como suporte à produção de software e como recurso pedagógico no ambiente de ensino.",
    ]
  ),

  // Projects
  ...sectionHeader("Projetos"),
  ...projectBlock(
    "Tatanea Temple — Plataforma Web Full Stack",
    "https://tatanea-temple.vercel.app",
    "Plataforma web completa para gerenciamento de fichas de personagem e compêndios de RPG. Destaque para a arquitetura de dados gerenciada remotamente, a complexidade da lógica front-end e o pipeline de deploy contínuo.",
    "JavaScript, HTML, CSS, Supabase (PostgreSQL), Vercel"
  ),
  ...projectBlock(
    "Sistema de Gerenciamento de Biblioteca",
    "https://github.com/TonhoDevi/SGBD---Java-v1.3",
    "Aplicação Java completa com arquitetura em camadas: CRUD de livros e categorias, validações, serviços, DAOs e integração com banco de dados relacional via ORM.",
    "Java, PostgreSQL, JPA/Hibernate"
  ),
  ...projectBlock(
    "Aplicativo de Agendamento de Turnos — BOPE",
    null,
    "Desenvolvimento de funcionalidades de backend, integração com PostgreSQL, consumo de APIs no frontend e participação na construção do fluxo principal de agendamentos. Projeto real em ambiente profissional.",
    "Java, Spring Boot, PostgreSQL, REST APIs"
  ),
  ...projectBlock(
    "Calculadora de Planos de Corte — Engenharia Civil",
    null,
    "Ferramenta web para cálculo otimizado de planos de corte de vigas metálicas, combinando lógica de engenharia estrutural com interface intuitiva para uso profissional.",
    "JavaScript, HTML, CSS"
  ),
  ...projectBlock(
    "Tales of the Dungeons Inquisitors — Jogo 2D",
    "https://tonhodevi.itch.io/tales-of-the-dungeos-inquisidors",
    "Jogo 2D completo com mecânicas de combate, exploração, sistema de colisão, HUD, animações, design de fases e publicação em plataforma pública.",
    "Godot Engine, GDScript"
  ),
];

// ─── Two-column layout via Table ──────────────────────────
const twoCol = new Table({
  width: { size: PAGE_W, type: WidthType.DXA },
  columnWidths: [COL_LEFT, COL_RIGHT],
  borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
  rows: [
    new TableRow({
      children: [
        // LEFT — sidebar
        new TableCell({
          width: { size: COL_LEFT, type: WidthType.DXA },
          shading: { fill: "1E40AF", type: ShadingType.CLEAR },
          borders: noBorders,
          margins: { top: 720, bottom: 720, left: 500, right: 400 },
          verticalAlign: VerticalAlign.TOP,
          children: sidebarChildren
        }),
        // RIGHT — main
        new TableCell({
          width: { size: COL_RIGHT, type: WidthType.DXA },
          borders: noBorders,
          margins: { top: 720, bottom: 720, left: 500, right: 500 },
          verticalAlign: VerticalAlign.TOP,
          children: mainChildren
        })
      ]
    })
  ]
});

// ─── Document ─────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "–",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 200 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: S_BODY } } },
    paragraphStyles: [
      { id: "Hyperlink", name: "Hyperlink", basedOn: "Normal",
        run: { color: ACCENT, underline: { type: "single" } } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MAR_T, bottom: MAR_B, left: MAR_L, right: MAR_R }
      }
    },
    children: [twoCol]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("./curriculo_antonio_henrique.docx", buf);
  console.log("Done!");
});