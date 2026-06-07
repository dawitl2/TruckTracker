import formidable from "formidable";
import * as XLSX from "xlsx";

const TARGET_LICENSE_PLATES = ["A06725/32431", "A09321/32699"];
const SUPABASE_URL = "https://ceaznmvgerreomiklcwo.supabase.co";
const SUPABASE_KEY = "sb_publishable_kF30JdMpqmsM9VmXPZLYAw_i8V58YJJ";
const SUPABASE_TABLE = "truck_arrivals";

function normalizePlate(value) {
  return String(value || "").trim().toUpperCase();
}

function formatDate(date) {
  return new Intl.DateTimeFormat("en-CA").format(date);
}

function formatTime(date) {
  return new Intl.DateTimeFormat("en-GB", {
    hour: "2-digit", minute: "2-digit", hour12: false,
  }).format(date);
}

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const form = formidable({});

  form.parse(req, async (err, _fields, files) => {
    if (err) return res.status(400).json({ error: "Could not parse file" });

    const file = files.file?.[0];
    if (!file) return res.status(400).json({ error: "No file received" });

    let workbook;
    try {
      workbook = XLSX.readFile(file.filepath);
    } catch {
      return res.status(400).json({ error: "Could not read Excel file" });
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    const now = new Date();
    const arrival_date = formatDate(now);
    const batch_time = formatTime(now);

    const targetRows = [];
    for (const row of matrix) {
      const serial = String(row[0] || "").trim();
      const plate = String(row[1] || "").trim();
      const code = String(row[2] || "").trim();
      if (!/^\d+$/.test(serial) || !plate || !code) continue;
      if (TARGET_LICENSE_PLATES.some((t) => normalizePlate(t) === normalizePlate(plate))) {
        targetRows.push({
          arrival_date,
          batch_time,
          license_plate: plate,
          arrival_code: code,
          product_type: String(row[3] || "").trim() || null,
          company: String(row[4] || "").trim() || null,
        });
      }
    }

    // No matches — redirect with not_found status so the app shows the right UI immediately
    if (!targetRows.length) {
      return res.redirect(302, `/?from=shortcut&status=not_found`);
    }

    // Save to Supabase
    let savedData;
    try {
      const response = await fetch(`${SUPABASE_URL}/rest/v1/${SUPABASE_TABLE}`, {
        method: "POST",
        headers: {
          apikey: SUPABASE_KEY,
          Authorization: `Bearer ${SUPABASE_KEY}`,
          "Content-Type": "application/json",
          Prefer: "return=representation",
        },
        body: JSON.stringify(targetRows),
      });

      if (!response.ok) {
        const errText = await response.text();
        console.error("Supabase error:", errText);
        return res.redirect(302, `/?from=shortcut&status=error`);
      }

      savedData = await response.json();
    } catch (fetchErr) {
      console.error("Fetch error:", fetchErr);
      return res.redirect(302, `/?from=shortcut&status=error`);
    }

    // Build redirect with everything the app needs to show the popup immediately —
    // no DB roundtrip required on the client side.
    const ids = savedData.map((r) => r.id).join(",");
    const plates = savedData.map((r) => encodeURIComponent(r.license_plate)).join(",");
    const codes = savedData.map((r) => encodeURIComponent(r.arrival_code)).join(",");
    const dates = savedData.map((r) => encodeURIComponent(r.arrival_date || "")).join(",");
    const times = savedData.map((r) => encodeURIComponent(r.batch_time || "")).join(",");

    return res.redirect(
      302,
      `/?from=shortcut&status=saved&ids=${ids}&plates=${plates}&codes=${codes}&dates=${dates}&times=${times}`
    );
  });
}