# Casco Digital - Site com Formulário de Contato

Site institucional da Casco Digital com formulário de contato funcional usando Cloudflare Pages Functions.

## Estrutura de Arquivos

```
casco-digital-site/
├── index.html              # Página principal do site
├── functions/
│   └── api/
│       └── contact.js      # Endpoint POST /api/contact (serverless)
├── wrangler.toml           # Config do Cloudflare
└── README.md               # Este arquivo
```

## Funcionalidades

✓ Design responsivo moderno
✓ Navbar com links rápidos (Suporte Remoto, Portal de Chamados)
✓ Formulário de contato funcional com validação
✓ Backend serverless (Cloudflare Pages Functions)
✓ Envio de email via API HTTP (Resend, SMTP-as-API, etc.)

## Setup - Cloudflare Pages (GRÁTIS)

### 1. Preparar o repositório

Crie um repositório Git com esta estrutura:

```
meu-repo/
├── index.html
└── functions/
    └── api/
        └── contact.js
```

Coloque o arquivo `index.html` na raiz e `contact.js` dentro de `functions/api/`.

### 2. Criar projeto no Cloudflare Pages

1. Acesse https://dash.cloudflare.com/ e vá em **Pages**
2. Clique em **Create a project** > **Connect to Git**
3. Conecte seu repositório (GitHub/GitLab)
4. Configure:
   - **Build command**: deixe vazio
   - **Build output directory**: `/`
   - **Root directory**: `/`

### 3. Configurar variáveis de ambiente

No dashboard do projeto Cloudflare Pages, vá em **Settings** > **Environment variables** e adicione:

| Variável        | Valor (exemplo)                      |
|----------------|--------------------------------------|
| `EMAIL_API_URL` | `https://api.resend.com/emails`     |
| `EMAIL_API_KEY` | Sua chave de API do provedor        |
| `EMAIL_FROM`   | `no-reply@cascodigital.com.br`      |
| `EMAIL_TO`     | `suporte@cascodigital.com.br`       |

### 4. Escolher provedor de email

O endpoint `functions/api/contact.js` usa uma API HTTP genérica para enviar emails.

#### Opção 1: Resend (Recomendado)

1. Crie conta grátis em https://resend.com
2. Gere uma API Key
3. Configure:
   - `EMAIL_API_URL`: `https://api.resend.com/emails`
   - `EMAIL_API_KEY`: Sua chave
   - `EMAIL_FROM`: Email verificado no Resend
   - `EMAIL_TO`: Seu email de destino

Documentação: https://resend.com/docs/send-with-cloudflare-workers

#### Opção 2: Outro provedor SMTP-as-API

Ajuste o código em `functions/api/contact.js` conforme a API do provedor escolhido:
- Mailgun
- SendGrid
- Postmark
- Etc.

### 5. Deploy

Faça push para o repositório Git. Cloudflare Pages vai automaticamente:
- Fazer deploy do `index.html`
- Expor a função em `/api/contact`

O site estará disponível em `https://seu-projeto.pages.dev`

## Desenvolvimento Local (Opcional)

### Requisitos
- Node.js 18+
- npm ou pnpm

### Instalar Wrangler
```bash
npm install -g wrangler
```

### Configurar variaveis locais
Edite `wrangler.toml` com suas credenciais de desenvolvimento.

### Rodar localmente
```bash
wrangler pages dev .
```

Acesse http://localhost:8788

## Personalização

### Alterar emails de destino

Edite a variável `EMAIL_TO` no Cloudflare Pages dashboard.

### Adicionar mais campos no formulário

1. Adicione os campos no HTML (`index.html`)
2. Capture-os no JavaScript (já está usando `FormData`)
3. Ajuste o template de email em `functions/api/contact.js`

### Mudar o estilo

Edite o CSS inline no `<style>` de `index.html`.

## Suporte

- Cloudflare Pages: https://developers.cloudflare.com/pages
- Resend: https://resend.com/docs
- Problemas: verifique os logs no dashboard Cloudflare Pages > Functions

## Tecnologias

- HTML5, CSS3, JavaScript (ES6+)
- Cloudflare Pages Functions (serverless)
- API HTTP para envio de email

---

**Casco Digital**
suporte@cascodigital.com.br
