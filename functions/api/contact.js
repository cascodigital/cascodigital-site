export async function onRequestPost(context) {
  try {
    const { request, env } = context;

    // Le body JSON
    const body = await request.json();
    const { name, email, phone, message } = body;

    if (!name || !email || !message) {
      return new Response(
        JSON.stringify({ ok: false, error: "Campos obrigatorios faltando." }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    // ===== INTEGRACAO COM PROVEDOR DE EMAIL =====
    // Este exemplo usa uma API HTTP generica
    // Configure as variaveis de ambiente no Cloudflare Pages:
    // - EMAIL_API_URL (ex: https://api.resend.com/emails)
    // - EMAIL_API_KEY (sua chave de API)
    // - EMAIL_FROM (ex: no-reply@cascodigital.com.br)
    // - EMAIL_TO (ex: suporte@cascodigital.com.br)

    const emailPayload = {
      from: env.EMAIL_FROM,
      to: env.EMAIL_TO,
      subject: "Novo contato - Site Casco Digital",
      text: `
Nome: ${name}
Email: ${email}
Telefone: ${phone || "-"}

Mensagem:
${message}
      `.trim(),
    };

    const sendResponse = await fetch(env.EMAIL_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${env.EMAIL_API_KEY}`,
      },
      body: JSON.stringify(emailPayload),
    });

    if (!sendResponse.ok) {
      const errorText = await sendResponse.text();
      console.error("Erro ao enviar email:", errorText);
      return new Response(
        JSON.stringify({ ok: false, error: "Falha ao enviar email." }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
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
