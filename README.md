# Casco Digital — Site Institucional

![Status](https://img.shields.io/badge/Status-Active-brightgreen)
![License](https://img.shields.io/badge/License-MIT-blue)
![Author](https://img.shields.io/badge/Author-Casco%20Digital-orange)

![Cloudflare](https://img.shields.io/badge/Cloudflare-Pages-F38020?style=flat-square&logo=cloudflare&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-CSS3-E34F26?style=flat-square&logo=html5&logoColor=white)
![Resend](https://img.shields.io/badge/Email-Resend-00D9FF?style=flat-square)

Site institucional da Casco Digital com formulario de contato funcional. Deploy automatico via Cloudflare Pages, envio de emails via Resend.

**Producao:** [cascodigital.com.br](https://cascodigital.com.br)

## Estrutura

```
cascodigital-site/
├── index.html              # Pagina principal
├── assets/images/          # Imagens do site
├── functions/api/
│   └── contact.js          # Serverless function (Cloudflare Pages)
└── README.md
```

## Formulario de Contato

O endpoint `POST /api/contact` (Cloudflare Pages Function) envia dois emails via Resend:
- Notificacao interna para a Casco Digital
- Confirmacao automatica para o cliente

## Deploy

### Cloudflare Pages

1. Conecte este repositorio em **Cloudflare Pages** > **Create a project** > **Connect to Git**
2. Configuracao de build:
   - Framework preset: **None**
   - Build command: (vazio)
   - Build output directory: `/`

### Variaveis de Ambiente

Configure em **Settings** > **Variables and Secrets** > **Production**:

| Variavel | Tipo | Valor |
|----------|------|-------|
| `EMAIL_API_KEY` | Secret | API Key do Resend (`re_...`) |
| `EMAIL_API_URL` | Text | `https://api.resend.com/emails` |
| `EMAIL_FROM` | Text | Remetente verificado no Resend |
| `EMAIL_TO` | Text | Email que recebe os contatos |

Apos configurar, faca um **Retry deployment** no ultimo deploy.

## Personalizacao

- **Destinatario:** altere `EMAIL_TO` nas variaveis de ambiente
- **Templates de email:** edite `internalPayload` e `clientPayload` em `functions/api/contact.js`

---

Desenvolvido com 🐢 (e cafe) por **Casco Digital**.
