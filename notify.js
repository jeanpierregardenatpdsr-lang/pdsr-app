// Fichier : api/notify.js — à placer dans le dossier /api à la racine du projet Vercel.
// Envoi d'un e-mail via Resend à chaque événement déclaré dans l'app PDSR.
// La clé reste côté serveur : elle n'est jamais visible dans le navigateur.
// Recommandé : définir RESEND_API_KEY dans Vercel (Settings > Environment Variables)
// et supprimer la clé en clair ci-dessous si le dépôt Git est public.

const RESEND_KEY = process.env.RESEND_API_KEY || "re_EXdpopAG_2HTXQXYuKcpJmRRf6Lq4V2Tz";

// Destinataires des alertes — à ajuster.
const DESTINATAIRES = ["jeanpierregardenatpdsr@gmail.com", "lmarcille1962@gmail.com", "omarngom21@yahoo.com"];

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "POST uniquement" });

  try {
    const b = req.body || {};
    const badge = b.eig ? "EIG — SIGNALEMENT" : (b.gravite === "Grave" ? "GRAVE" : b.gravite === "Moyen" ? "Moyen" : "Léger");
    const subject = "[PDSR" + (b.site ? " " + b.site : "") + "] " + badge + " — " + (b.titre || "Événement") + (b.jeune ? " (" + b.jeune + ")" : "");
    const esc = (s) => String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br/>");
    const couleur = b.eig || b.gravite === "Grave" ? "#C62828" : b.gravite === "Moyen" ? "#E65100" : "#F9A825";
    const html =
      '<div style="font-family:Arial,sans-serif;max-width:600px">' +
      '<div style="background:' + couleur + ';color:#fff;padding:14px 18px;border-radius:10px 10px 0 0;font-size:16px;font-weight:bold">' + esc(badge) + " — " + esc(b.titre) + "</div>" +
      '<div style="border:1px solid #ddd;border-top:none;padding:18px;border-radius:0 0 10px 10px">' +
      "<p><b>Jeune :</b> " + esc(b.jeune || "—") + "<br/>" +
      "<b>Site :</b> " + esc(b.site || "—") + "<br/>" +
      "<b>Date :</b> " + esc(b.date || "—") + "<br/>" +
      "<b>Gravité :</b> " + esc(b.gravite || "—") + (b.eig ? " · <b style='color:#C62828'>EIG art. L331-8-1 CASF</b>" : "") + "<br/>" +
      "<b>Déclaré par :</b> " + esc(b.author || "—") + "</p>" +
      '<p style="background:#f7f5f0;padding:12px;border-radius:8px">' + esc(b.description || "(sans description)") + "</p>" +
      '<p style="font-size:12px;color:#888">Notification automatique PDSR — pdsr-app.vercel.app</p>' +
      "</div></div>";

    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: "Bearer " + RESEND_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ from: "PDSR App <onboarding@resend.dev>", to: DESTINATAIRES, subject, html })
    });
    const j = await r.json();
    return res.status(r.ok ? 200 : 502).json(j);
  } catch (e) {
    return res.status(500).json({ error: String(e) });
  }
};
