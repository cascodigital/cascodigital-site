/**
 * Cloudflare Pages Function - The Guardian
 * Endpoint GET /api/audit?domain=example.com
 * Audita SPF, DMARC e MX via DNS over HTTPS (Cloudflare DoH)
 * e retorna score de risco 0-100.
 *
 * Autor: Andre Kittler / Casco Digital
 */

const DOH = "https://cloudflare-dns.com/dns-query";

async function dnsQuery(name, type) {
  const url = `${DOH}?name=${encodeURIComponent(name)}&type=${type}`;
  const resp = await fetch(url, { headers: { Accept: "application/dns-json" } });
  if (!resp.ok) return null;
  return resp.json();
}

function analyzeSPF(records) {
  const raw = records.find((r) => r.data && r.data.replace(/"/g, "").trimStart().startsWith("v=spf1"));
  if (!raw) return { exists: false, value: null, score: 40, status: "missing", label: "Ausente" };

  const val = raw.data.replace(/"/g, "");
  if (val.includes("+all"))
    return { exists: true, value: val, score: 35, status: "dangerous", label: "Perigoso (+all)" };
  if (val.includes("?all"))
    return { exists: true, value: val, score: 20, status: "weak", label: "Fraco (?all)" };
  if (val.includes("~all"))
    return { exists: true, value: val, score: 15, status: "softfail", label: "Softfail (~all)" };
  if (val.includes("-all"))
    return { exists: true, value: val, score: 5, status: "strict", label: "Restritivo (-all)" };

  return { exists: true, value: val, score: 15, status: "incomplete", label: "Incompleto" };
}

function analyzeDMARC(records) {
  const raw = records.find((r) => r.data && r.data.includes("v=DMARC1"));
  if (!raw) return { exists: false, value: null, score: 40, status: "missing", label: "Ausente" };

  const val = raw.data.replace(/"/g, "");
  if (val.includes("p=reject"))
    return { exists: true, value: val, score: 0, status: "strict", label: "Rejeição (p=reject)" };
  if (val.includes("p=quarantine"))
    return { exists: true, value: val, score: 5, status: "moderate", label: "Quarentena (p=quarantine)" };
  if (val.includes("p=none"))
    return { exists: true, value: val, score: 20, status: "monitoring", label: "Monitoramento (p=none)" };

  return { exists: true, value: val, score: 15, status: "incomplete", label: "Incompleto" };
}

export async function onRequestGet(context) {
  const { request } = context;
  const url = new URL(request.url);
  const rawDomain = url.searchParams.get("domain") || "";

  const domain = rawDomain
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "")
    .toLowerCase()
    .trim();

  if (!domain) {
    return json({ ok: false, error: "Parâmetro 'domain' obrigatório." }, 400);
  }

  try {
    const [txtRes, dmarcRes, mxRes] = await Promise.all([
      dnsQuery(domain, "TXT"),
      dnsQuery(`_dmarc.${domain}`, "TXT"),
      dnsQuery(domain, "MX"),
    ]);

    const spf = analyzeSPF(txtRes?.Answer || []);
    const dmarc = analyzeDMARC(dmarcRes?.Answer || []);

    const mxRecords = (mxRes?.Answer || []).map((r) => r.data).filter(Boolean);
    const mxExists = mxRecords.length > 0;
    const mxScore = mxExists ? 0 : 10;

    const score = Math.min(100, spf.score + dmarc.score + mxScore);

    return json({
      ok: true,
      domain,
      score,
      spf,
      dmarc,
      mx: { exists: mxExists, records: mxRecords, score: mxScore },
    });
  } catch (err) {
    return json({ ok: false, error: err.message }, 500);
  }
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*",
    },
  });
}
