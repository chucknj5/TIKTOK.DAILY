// TikTok W.E. Report — shared save for everyone on the Netlify link.
// GET  /api/state  -> latest saved data
// POST /api/state  -> save (also keeps one backup copy per day)
import { getStore } from "@netlify/blobs";

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { "content-type": "application/json", "cache-control": "no-store" } });

export default async (req) => {
  const store = getStore({ name: "tiktok-we-report", consistency: "strong" });

  if (req.method === "GET") {
    const data = await store.get("state", { type: "text" });
    return new Response(data || "{}", { headers: { "content-type": "application/json", "cache-control": "no-store" } });
  }

  if (req.method === "POST") {
    const body = await req.text();
    let d;
    try { d = JSON.parse(body); } catch { return json({ ok: false, error: "bad data" }, 400); }
    if (!d || d.app !== "tiktok-we") return json({ ok: false, error: "wrong app" }, 400);
    await store.set("state", body);
    await store.set("backup/" + new Date().toISOString().slice(0, 10), body);
    return json({ ok: true, savedAt: d.savedAt });
  }

  return json({ ok: false, error: "method not allowed" }, 405);
};

export const config = { path: "/api/state" };
