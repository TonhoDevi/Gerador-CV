# Gerador de Currículos

Gere currículos em formato Word (.docx) a partir de dados em JSON.

## Estrutura do Projeto

```
├── curriculo.json          # Exemplo de JSON com dados do currículo
├── main.js                 # Script Node.js para gerar o currículo
├── frontend/
│   ├── index.html          # Interface web
│   ├── style.css           # Estilos
│   └── app.js              # Lógica de geração no navegador
└── package.json
```

## Como Usar (Frontend - Navegador)

1. Abra `frontend/index.html` no navegador
2. Preencha o formulário com suas informações
3. Ou clique em **Importar JSON** para carregar um arquivo JSON
4. Clique em **Gerar Currículo DOCX** para baixar o arquivo

## Estrutura do JSON

```json
{
  "firstName": "Nome",
  "lastName": "Sobrenome",
  "jobTitle": "Cargo",
  "contact": {
    "address": "Cidade, UF",
    "phone": "(00) 00000-0000",
    "email": "email@exemplo.com",
    "github": "github.com/usuario",
    "linkedin": "linkedin.com/in/usuario"
  },
  "profile": "Descrição do perfil profissional",
  "technologies": ["Tech 1", "Tech 2"],
  "languages": [
    { "name": "Português", "level": "Nativo" }
  ],
  "skills": ["Habilidade 1", "Habilidade 2"],
  "education": [
    { "course": "Curso", "institution": "Instituição", "period": "2020 - 2024" }
  ],
  "experiences": [
    { "role": "Cargo", "company": "Empresa", "period": "2020 - Presente", "location": "Cidade, UF", "description": "Descrição das atividades" }
  ],
  "projects": [
    { "name": "Projeto", "url": "https://...", "description": "Descrição", "stack": "Tecnologias" }
  ]
}
```

## Como Usar (Node.js)

```bash
node main.js
```

Gera o arquivo `curriculo_antonio_henrique.docx` com os dados hardcoded.