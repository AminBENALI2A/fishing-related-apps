// npm install google-play-scraper xlsx

import gplay from "google-play-scraper";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const INPUT_FILE = "playstore_results.xlsx";
const OUTPUT_FILE = "./playstore_results_with_descriptions.xlsx";
const LINK_COLUMN = "Link";

// Read Excel
const wb_in = XLSX.readFile(INPUT_FILE);
const rows = XLSX.utils.sheet_to_json(wb_in.Sheets[wb_in.SheetNames[0]]);

function extractAppId(url) {
  try {
    return new URL(url.trim()).searchParams.get("id");
  } catch {
    return url.trim();
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const results = [];
for (let i = 0; i < rows.length; i++) {
  const row = rows[i];
  const appId = extractAppId(row[LINK_COLUMN]);
  console.log(`[${i + 1}/${rows.length}] ${appId}`);

  try {
    const app = await gplay.app({ appId, lang: "en", country: "us" });
    let reviews = [];
    try {
      reviews = (await gplay.reviews({ appId, lang: "en", num: 100 })).data;
    } catch {}

    results.push({
      ...row,
      "Scraped Name": app.title,
      Developer: app.developer,
      "Developer Email": app.developerEmail || "",
      "Developer Website": app.developerWebsite || "",
      "Category (Scraped)": app.genre,
      Rating: app.score,
      "Ratings Count": app.ratings,
      "Reviews Count": app.reviews,
      "5 Stars": app.histogram?.[5] || 0,
      "4 Stars": app.histogram?.[4] || 0,
      "3 Stars": app.histogram?.[3] || 0,
      "2 Stars": app.histogram?.[2] || 0,
      "1 Star": app.histogram?.[1] || 0,
      Installs: app.installs,
      "Min Installs": app.minInstalls,
      Free: app.free ? "Yes" : "No",
      Price: app.price,
      "In-App Purchases": app.offersIAP ? "Yes" : "No",
      "Ad Supported": app.adSupported ? "Yes" : "No",
      Version: app.version,
      "Android Required": app.androidVersion,
      "Last Updated": app.updated,
      Released: app.released,
      "Content Rating": app.contentRating,
      Description: app.description,
      Summary: app.summary,
      "Recent Changes": app.recentChanges || "",
      "Privacy Policy": app.privacyPolicy || "",
      "App URL": app.url,
      "Review 1 Author": reviews[0]?.userName || "",
      "Review 1 Score": reviews[0]?.score || "",
      "Review 1 Text": reviews[0]?.text || "",
      "Review 2 Author": reviews[1]?.userName || "",
      "Review 2 Score": reviews[1]?.score || "",
      "Review 2 Text": reviews[1]?.text || "",
      "Review 3 Author": reviews[2]?.userName || "",
      "Review 3 Score": reviews[2]?.score || "",
      "Review 3 Text": reviews[2]?.text || "",
      Status: "OK",
    });
    console.log(`  ✓ ${app.score} stars | ${app.installs} installs`);
  } catch (e) {
    results.push({ ...row, Status: "ERROR", Error: e.message });
    console.log(`  ✗ ${e.message}`);
  }

  await sleep(1500);
}

// Write Excel
const wb_out = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(
  wb_out,
  XLSX.utils.json_to_sheet(results),
  "Results",
);
XLSX.writeFile(wb_out, OUTPUT_FILE);
console.log(`\nDone → ${OUTPUT_FILE}`);
