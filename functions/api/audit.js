const DOH = "https://cloudflare-dns.com/dns-query";

async function dnsQuery(name, type) {
  const url = `${DOH}?name=${encodeURIComponent(name)}&type=${type}`;
  const res = await fetch(url, { headers: { Accept: "application/dns-json" } });
  if (!res.ok) return [];
  const data = await res.json();
  return data.Answer || [];
}

function analyzeSPF(records) {
  const spfRecord = records.find(r => r.data && r.data.includes("v=spf1"));
  if (!spfRecord) return { status: "missing", label: "Ausente", value: null };

  const val = spfRecord.data.replace(/"/g, "").trim();

  if (val.includes("+all")) return { status: "dangerous",  label: "Perigoso",  value: val };
  if (val.includes("?all")) return { status: "weak",       label: "Fraco",     value: val };
  if (val.includes("~all")) return { status: "softfail",   label: "Softfail",  value: val };
  if (val.includes("-all")) return { status: "strict",     label: "Seguro",    value: val };
  if (val.includes("v=spf1")) return { status: "incomplete", label: "Incompleto", value: val };

  return { status: "missing", label: "Ausente", value: null };
}

function analyzeDMARC(records) {
  const dmarcRecord = records.find(r => r.data && r.data.includes("v=DMARC1"));
  if (!dmarcRecord) return { status: "missing", label: "Ausente", value: null };

  const val = dmarcRecord.data.replace(/"/g, "").trim();

  if (val.includes("p=reject"))     return { status: "strict",     label: "Máximo",      value: val };
  if (val.includes("p=quarantine")) return { status: "moderate",   label: "Moderado",    value: val };
  if (val.includes("p=none"))       return { status: "monitoring", label: "Monitorando", value: val };

  return { status: "incomplete", label: "Incompleto", value: val };
}

function calcScore(spf, dmarc, mx) {
  let score = 100;

  // SPF penalties
  if      (spf.status === "missing")    score -= 35;
  else if (spf.status === "dangerous")  score -= 40;
  else if (spf.status === "weak")       score -= 30;
  else if (spf.status === "softfail")   score -= 10;
  else if (spf.status === "incomplete") score -= 15;

  // DMARC penalties
  if      (dmarc.status === "missing")    score -= 40;
  else if (dmarc.status === "monitoring") score -= 20;
  else if (dmarc.status === "moderate")   score -= 5;
  else if (dmarc.status === "incomplete") score -= 15;

  // MX penalty (minor)
  if (!mx.exists) score -= 5;

  return Math.max(0, Math.min(100, score));
}

export async function onRequestGet(context) {
  const url = new URL(context.request.url);
  const domain = url.searchParams.get("domain");

  const cors = {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
  };

  if (!domain) {
    return new Response(JSON.stringify({ ok: false, error: "Domínio não informado." }), { status: 400, headers: cors });
  }

  // Sanitize
  const cleanDomain = domain.trim().toLowerCase().replace(/^https?:\/\//, "").replace(/\/.*$/, "");

  try {
    const [txtRecords, dmarcRecords, mxRecords] = await Promise.all([
      dnsQuery(cleanDomain, "TXT"),
      dnsQuery(`_dmarc.${cleanDomain}`, "TXT"),
      dnsQuery(cleanDomain, "MX"),
    ]);

    const spf   = analyzeSPF(txtRecords);
    const dmarc = analyzeDMARC(dmarcRecords);
    const mx    = {
      exists:  mxRecords.length > 0,
      records: mxRecords.map(r => r.data).slice(0, 5),
    };
    const score = calcScore(spf, dmarc, mx);

    return new Response(JSON.stringify({ ok: true, domain: cleanDomain, score, spf, dmarc, mx }), {
      status: 200,
      headers: cors,
    });
  } catch (err) {
    return new Response(JSON.stringify({ ok: false, error: "Erro ao consultar DNS. Tente novamente." }), {
      status: 500,
      headers: cors,
    });
  }
}
