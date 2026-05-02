const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  LevelFormat, ExternalHyperlink
} = docx;

const ACCENT = "1A56DB";
const DARK = "1E293B";
const MID = "475569";
const WHITE = "FFFFFF";
const RULE = "BFDBFE";

const S_NAME = 28;
const S_TITLE = 22;
const S_SECTION = 22;
const S_BODY = 18;
const S_SMALL = 16;

const PAGE_W = 11906;
const PAGE_H = 16838;
const MAR_T = 720;
const MAR_B = 720;
const MAR_L = 0;
const MAR_R = 0;
const COL_LEFT = 2900;
const COL_RIGHT = 9006;

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function rule(color = RULE) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color, space: 1 } },
    children: []
  });
}

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

function sideText(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text, size: S_SMALL, color: WHITE, font: "Arial", ...opts })
    ]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text, size: S_BODY, color: DARK, font: "Arial" })
    ]
  });
}

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

function addLanguage() {
  const container = document.getElementById('languagesContainer');
  const entry = document.createElement('div');
  entry.className = 'language-entry';
  entry.innerHTML = `
    <input type="text" class="lang-name" placeholder="Idioma (ex: Inglês)">
    <input type="text" class="lang-level" placeholder="Nível (ex: Avançado)">
    <button type="button" class="btn-remove" onclick="removeEntry(this)">✕</button>
  `;
  container.appendChild(entry);
}

function addEducation() {
  const container = document.getElementById('educationContainer');
  const entry = document.createElement('div');
  entry.className = 'education-entry';
  entry.innerHTML = `
    <input type="text" class="edu-course" placeholder="Curso">
    <input type="text" class="edu-institution" placeholder="Instituição">
    <input type="text" class="edu-period" placeholder="Período (ex: 2023 – 2025)">
    <button type="button" class="btn-remove" onclick="removeEntry(this)">✕</button>
  `;
  container.appendChild(entry);
}

function addExperience() {
  const container = document.getElementById('experienceContainer');
  const entry = document.createElement('div');
  entry.className = 'experience-entry';
  entry.innerHTML = `
    <input type="text" class="exp-role" placeholder="Cargo">
    <input type="text" class="exp-company" placeholder="Empresa">
    <div class="form-row">
      <input type="text" class="exp-period" placeholder="Período (ex: Jan 2023 – Presente)">
      <input type="text" class="exp-location" placeholder="Local (ex: Recife, PE)">
    </div>
    <textarea class="exp-description" rows="3" placeholder="Liste suas atividades e realizações (uma por linha)"></textarea>
    <button type="button" class="btn-remove" onclick="removeEntry(this)">✕ Remover</button>
  `;
  container.appendChild(entry);
}

function addProject() {
  const container = document.getElementById('projectsContainer');
  const entry = document.createElement('div');
  entry.className = 'project-entry';
  entry.innerHTML = `
    <input type="text" class="proj-name" placeholder="Nome do Projeto">
    <input type="text" class="proj-url" placeholder="URL (opcional)">
    <textarea class="proj-description" rows="2" placeholder="Descrição do projeto"></textarea>
    <input type="text" class="proj-stack" placeholder="Stack/Tecnologias">
    <button type="button" class="btn-remove" onclick="removeEntry(this)">✕ Remover</button>
  `;
  container.appendChild(entry);
}

function removeEntry(btn) {
  btn.parentElement.remove();
}

function getFormData() {
  const firstName = document.getElementById('firstName').value.trim();
  const lastName = document.getElementById('lastName').value.trim();
  const jobTitle = document.getElementById('jobTitle').value.trim();
  const address = document.getElementById('address').value.trim();
  const phone = document.getElementById('phone').value.trim();
  const email = document.getElementById('email').value.trim();
  const github = document.getElementById('github').value.trim();
  const linkedin = document.getElementById('linkedin').value.trim();
  const profile = document.getElementById('profile').value.trim();
  const technologies = document.getElementById('technologies').value.trim();
  const skills = document.getElementById('skills').value.trim();

  const languages = [];
  document.querySelectorAll('.language-entry').forEach(entry => {
    const name = entry.querySelector('.lang-name').value.trim();
    const level = entry.querySelector('.lang-level').value.trim();
    if (name) languages.push({ name, level });
  });

  const education = [];
  document.querySelectorAll('.education-entry').forEach(entry => {
    const course = entry.querySelector('.edu-course').value.trim();
    const institution = entry.querySelector('.edu-institution').value.trim();
    const period = entry.querySelector('.edu-period').value.trim();
    if (course) education.push({ course, institution, period });
  });

  const experiences = [];
  document.querySelectorAll('.experience-entry').forEach(entry => {
    const role = entry.querySelector('.exp-role').value.trim();
    const company = entry.querySelector('.exp-company').value.trim();
    const period = entry.querySelector('.exp-period').value.trim();
    const location = entry.querySelector('.exp-location').value.trim();
    const description = entry.querySelector('.exp-description').value.trim();
    if (role) experiences.push({ role, company, period, location, description });
  });

  const projects = [];
  document.querySelectorAll('.project-entry').forEach(entry => {
    const name = entry.querySelector('.proj-name').value.trim();
    const url = entry.querySelector('.proj-url').value.trim();
    const description = entry.querySelector('.proj-description').value.trim();
    const stack = entry.querySelector('.proj-stack').value.trim();
    if (name) projects.push({ name, url, description, stack });
  });

  return {
    firstName, lastName, jobTitle, address, phone, email, github, linkedin,
    profile, technologies, skills, languages, education, experiences, projects
  };
}

function populateForm(data) {
  document.getElementById('firstName').value = data.firstName || '';
  document.getElementById('lastName').value = data.lastName || '';
  document.getElementById('jobTitle').value = data.jobTitle || '';

  if (data.contact) {
    document.getElementById('address').value = data.contact.address || '';
    document.getElementById('phone').value = data.contact.phone || '';
    document.getElementById('email').value = data.contact.email || '';
    document.getElementById('github').value = data.contact.github || '';
    document.getElementById('linkedin').value = data.contact.linkedin || '';
  }

  document.getElementById('profile').value = data.profile || '';

  if (data.technologies && Array.isArray(data.technologies)) {
    document.getElementById('technologies').value = data.technologies.join('\n');
  }

  const langContainer = document.getElementById('languagesContainer');
  langContainer.innerHTML = '';
  (data.languages || []).forEach(lang => {
    const entry = document.createElement('div');
    entry.className = 'language-entry';
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.className = 'lang-name';
    nameInput.placeholder = 'Idioma';
    nameInput.value = lang.name || '';
    const levelInput = document.createElement('input');
    levelInput.type = 'text';
    levelInput.className = 'lang-level';
    levelInput.placeholder = 'Nível';
    levelInput.value = lang.level || '';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕';
    removeBtn.onclick = function() { removeEntry(this); };
    entry.appendChild(nameInput);
    entry.appendChild(levelInput);
    entry.appendChild(removeBtn);
    langContainer.appendChild(entry);
  });
  if (langContainer.children.length === 0) {
    addLanguage();
  }

  document.getElementById('skills').value = (data.skills || []).join('\n');

  const eduContainer = document.getElementById('educationContainer');
  eduContainer.innerHTML = '';
  (data.education || []).forEach(edu => {
    const entry = document.createElement('div');
    entry.className = 'education-entry';
    const courseInput = document.createElement('input');
    courseInput.type = 'text';
    courseInput.className = 'edu-course';
    courseInput.placeholder = 'Curso';
    courseInput.value = edu.course || '';
    const instInput = document.createElement('input');
    instInput.type = 'text';
    instInput.className = 'edu-institution';
    instInput.placeholder = 'Instituição';
    instInput.value = edu.institution || '';
    const periodInput = document.createElement('input');
    periodInput.type = 'text';
    periodInput.className = 'edu-period';
    periodInput.placeholder = 'Período';
    periodInput.value = edu.period || '';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕';
    removeBtn.onclick = function() { removeEntry(this); };
    entry.appendChild(courseInput);
    entry.appendChild(instInput);
    entry.appendChild(periodInput);
    entry.appendChild(removeBtn);
    eduContainer.appendChild(entry);
  });
  if (eduContainer.children.length === 0) {
    addEducation();
  }

  const expContainer = document.getElementById('experienceContainer');
  expContainer.innerHTML = '';
  (data.experiences || []).forEach(exp => {
    const entry = document.createElement('div');
    entry.className = 'experience-entry';
    const roleInput = document.createElement('input');
    roleInput.type = 'text';
    roleInput.className = 'exp-role';
    roleInput.placeholder = 'Cargo';
    roleInput.value = exp.role || '';
    const companyInput = document.createElement('input');
    companyInput.type = 'text';
    companyInput.className = 'exp-company';
    companyInput.placeholder = 'Empresa';
    companyInput.value = exp.company || '';
    const formRow = document.createElement('div');
    formRow.className = 'form-row';
    const periodInput = document.createElement('input');
    periodInput.type = 'text';
    periodInput.className = 'exp-period';
    periodInput.placeholder = 'Período';
    periodInput.value = exp.period || '';
    const locationInput = document.createElement('input');
    locationInput.type = 'text';
    locationInput.className = 'exp-location';
    locationInput.placeholder = 'Local';
    locationInput.value = exp.location || '';
    formRow.appendChild(periodInput);
    formRow.appendChild(locationInput);
    const descTextarea = document.createElement('textarea');
    descTextarea.className = 'exp-description';
    descTextarea.rows = 3;
    descTextarea.value = exp.description || '';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕ Remover';
    removeBtn.onclick = function() { removeEntry(this); };
    entry.appendChild(roleInput);
    entry.appendChild(companyInput);
    entry.appendChild(formRow);
    entry.appendChild(descTextarea);
    entry.appendChild(removeBtn);
    expContainer.appendChild(entry);
  });
  if (expContainer.children.length === 0) {
    addExperience();
  }

  const projContainer = document.getElementById('projectsContainer');
  projContainer.innerHTML = '';
  (data.projects || []).forEach(proj => {
    const entry = document.createElement('div');
    entry.className = 'project-entry';
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.className = 'proj-name';
    nameInput.placeholder = 'Nome do Projeto';
    nameInput.value = proj.name || '';
    const urlInput = document.createElement('input');
    urlInput.type = 'text';
    urlInput.className = 'proj-url';
    urlInput.placeholder = 'URL (opcional)';
    urlInput.value = proj.url || '';
    const descTextarea = document.createElement('textarea');
    descTextarea.className = 'proj-description';
    descTextarea.rows = 2;
    descTextarea.value = proj.description || '';
    const stackInput = document.createElement('input');
    stackInput.type = 'text';
    stackInput.className = 'proj-stack';
    stackInput.placeholder = 'Stack/Tecnologias';
    stackInput.value = proj.stack || '';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕ Remover';
    removeBtn.onclick = function() { removeEntry(this); };
    entry.appendChild(nameInput);
    entry.appendChild(urlInput);
    entry.appendChild(descTextarea);
    entry.appendChild(stackInput);
    entry.appendChild(removeBtn);
    projContainer.appendChild(entry);
  });
  if (projContainer.children.length === 0) {
    addProject();
  }
}

function loadJSON() {
  document.getElementById('jsonFileInput').click();
}

document.getElementById('jsonFileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(event) {
    try {
      const text = event.target.result;
      const data = JSON.parse(text);
      console.log('JSON carregado:', data);
      populateForm(data);
      alert('JSON importado com sucesso!');
    } catch (err) {
      console.error('Erro ao parsear JSON:', err);
      alert('Erro ao ler JSON: ' + err.message);
    }
  };
  reader.onerror = function() {
    console.error('Erro ao ler arquivo:', reader.error);
    alert('Erro ao ler arquivo.');
  };
  reader.readAsText(file);
  this.value = '';
});

async function generateCV() {
  const data = getFormData();

  if (!data.firstName || !data.lastName) {
    alert('Por favor, preencha pelo menos seu nome e sobrenome.');
    return;
  }

  const sidebarChildren = [
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: data.firstName.toUpperCase(), bold: true, size: S_NAME, color: WHITE, font: "Arial" })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 20 },
      children: [
        new TextRun({ text: data.lastName.toUpperCase(), bold: true, size: S_NAME - 6, color: "BFDBFE", font: "Arial" })
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 300 },
      children: [
        new TextRun({ text: data.jobTitle || "Profissão", size: S_TITLE, color: "93C5FD", font: "Arial", italics: true })
      ]
    }),
    ...sideHeader("Contato"),
  ];

  if (data.address) sidebarChildren.push(sideText(`📍 ${data.address}`));
  if (data.phone) sidebarChildren.push(sideText(`📞 ${data.phone}`));
  if (data.email) sidebarChildren.push(sideText(`✉ ${data.email}`));
  if (data.github) sidebarChildren.push(sideText(`🔗 ${data.github}`));
  if (data.linkedin) sidebarChildren.push(sideText(`🔗 ${data.linkedin}`));

  if (data.technologies) {
    sidebarChildren.push(...sideHeader("Tecnologias"));
    const techLines = data.technologies.split('\n').filter(t => t.trim());
    techLines.forEach(tech => {
      sidebarChildren.push(sideText(tech));
    });
  }

  if (data.languages.length > 0) {
    sidebarChildren.push(...sideHeader("Idiomas"));
    data.languages.forEach((lang, idx) => {
      if (idx > 0) {
        sidebarChildren.push(new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }));
      }
      sidebarChildren.push(sideText(lang.name, { bold: true }));
      sidebarChildren.push(sideText(lang.level));
    });
  }

  if (data.skills) {
    sidebarChildren.push(...sideHeader("Competências"));
    const skillLines = data.skills.split('\n').filter(s => s.trim());
    skillLines.forEach(skill => {
      sidebarChildren.push(sideText(`• ${skill}`));
    });
  }

  if (data.education.length > 0) {
    sidebarChildren.push(...sideHeader("Formação"));
    data.education.forEach((edu, idx) => {
      if (idx > 0) {
        sidebarChildren.push(new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }));
      }
      sidebarChildren.push(sideText(edu.course, { bold: true }));
      sidebarChildren.push(sideText(`${edu.institution} · ${edu.period}`));
    });
  }

  const mainChildren = [];

  if (data.profile) {
    mainChildren.push(...sectionHeader("Perfil Profissional"));
    mainChildren.push(new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [new TextRun({ text: data.profile, size: S_BODY, color: DARK, font: "Arial" })]
    }));
  }

  if (data.experiences.length > 0) {
    mainChildren.push(...sectionHeader("Experiência"));
    data.experiences.forEach(exp => {
      const bullets = exp.description ? exp.description.split('\n').filter(b => b.trim()) : [];
      mainChildren.push(...expBlock(exp.role, exp.company, exp.period, exp.location, bullets));
    });
  }

  if (data.projects.length > 0) {
    mainChildren.push(...sectionHeader("Projetos"));
    data.projects.forEach(proj => {
      mainChildren.push(...projectBlock(proj.name, proj.url || null, proj.description, proj.stack));
    });
  }

  const twoCol = new Table({
    width: { size: PAGE_W, type: WidthType.DXA },
    columnWidths: [COL_LEFT, COL_RIGHT],
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: COL_LEFT, type: WidthType.DXA },
            shading: { fill: "1E40AF", type: ShadingType.CLEAR },
            borders: noBorders,
            margins: { top: 720, bottom: 720, left: 500, right: 400 },
            verticalAlign: VerticalAlign.TOP,
            children: sidebarChildren
          }),
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

  const doc = new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "–",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 200 } } }
        }]
      }]
    },
    styles: {
      default: { document: { run: { font: "Arial", size: S_BODY } } },
      paragraphStyles: [{
        id: "Hyperlink", name: "Hyperlink", basedOn: "Normal",
        run: { color: ACCENT, underline: { type: "single" } }
      }]
    },
    sections: [{
      properties: {
        page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MAR_T, bottom: MAR_B, left: MAR_L, right: MAR_R } }
      },
      children: [twoCol]
    }]
  });

  const blob = await Packer.toBlob(doc);
  const filename = `curriculo_${data.firstName.toLowerCase()}_${data.lastName.toLowerCase()}.docx`;
  saveAs(blob, filename);
}

document.getElementById('cvForm').addEventListener('submit', function(e) {
  e.preventDefault();
  generateCV();
});