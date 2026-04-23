/**
 * Cloudflare Pages Function - Formulario de Contato
 *
 * Endpoint POST /api/contact
 * Recebe dados do formulario (name, email, phone, message) e envia dois emails
 * via Resend: notificacao interna + confirmacao para o cliente.
 *
 * Env vars: EMAIL_FROM, EMAIL_TO, EMAIL_API_URL, EMAIL_API_KEY
 *
 * @param {Object} context - Cloudflare Pages Function context (request, env)
 * @returns {Response} JSON { ok: true/false, error?: string }
 *
 * Autor: Andre Kittler / Casco Digital
 */
export async function onRequestPost(context) {
  try {
    const { request, env } = context;

    const body = await request.json();
    const { name, email, phone, message } = body;

    if (!name || !email || !message) {
      return new Response(
        JSON.stringify({ ok: false, error: "Campos obrigatorios faltando." }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    // EMAIL 1: Para a Casco Digital (interno)
    const internalPayload = {
      from: env.EMAIL_FROM,
      to: env.EMAIL_TO,
      subject: "Novo contato - Site Casco Digital",
      text: `
Novo contato recebido pelo site Casco Digital.

Nome: ${name}
Email: ${email}
Telefone: ${phone || "-"}

Mensagem:
${message}
      `.trim(),
    };

    const internalResp = await fetch(env.EMAIL_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${env.EMAIL_API_KEY}`,
      },
      body: JSON.stringify(internalPayload),
    });

    if (!internalResp.ok) {
      const errText = await internalResp.text();
      console.error("Erro ao enviar email interno:", errText);
      return new Response(
        JSON.stringify({ ok: false, error: "Falha ao notificar equipe." }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }

    // EMAIL 2: Confirmacao para o cliente (externo)
    const clientPayload = {
      from: env.EMAIL_FROM,
      to: email,
      subject: "Recebemos sua mensagem - Casco Digital",
      text: `
Olá ${name},

Recebemos sua mensagem enviada pelo site da Casco Digital e já estamos analisando.

Resumo do que você enviou:

Nome: ${name}
Email: ${email}
Telefone: ${phone || "-"}

Mensagem:
${message}

Em breve entraremos em contato.

Casco Digital
Consultoria em M365, PowerShell e Infraestrutura
      `.trim(),
    };

    const clientResp = await fetch(env.EMAIL_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${env.EMAIL_API_KEY}`,
      },
      body: JSON.stringify(clientPayload),
    });

    if (!clientResp.ok) {
      const errText = await clientResp.text();
      console.error("Erro ao enviar email para o cliente:", errText);
      // Nao quebra a requisicao para o front, so loga o erro
    }

    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error("Erro interno:", err);
    return new Response(
      JSON.stringify({ ok: false, error: "Erro interno do servidor." }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}
