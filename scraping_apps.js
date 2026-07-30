const gplay = require("google-play-scraper");
const ExcelJS = require("exceljs");
const path = require("path");

// ─── Search Terms ────────────────────────────────────────────────────────────
const SEARCH_TERMS = [
  "Fishing",
  "Oceanology",
  "Geographic Information System",
  "Maritime traffic",
  "Maritime navigation",
  "Oceanography",
  "Biodiversity",
  "Sea",
  "Climate change oceans",
  "Fishing app",
  "Fish finder",
  "Fishing forecast",
  "Fishing map",
  "Fishing log",
  "Fishing tracker",
  "Fishing spots",
  "Fly fishing",
  "Sea fishing",
  "Boat navigation",
  "Tide chart",
  "Weather marine",
  "Sonar fish",
];

const RESULTS_PER_TERM = 50; // max 250
const DELAY_MS = 2000;       // delay between requests (be polite)
const OUTPUT_FILE = "playstore_results.xlsx";

// ─── Helpers ─────────────────────────────────────────────────────────────────
const delay = (ms) => new Promise((res) => setTimeout(res, ms));

function formatDate(timestamp) {
  if (!timestamp) return "N/A";
  try {
    return new Date(timestamp).toLocaleDateString("en-GB", {
      day: "2-digit", month: "short", year: "numeric"
    });
  } catch {
    return "N/A";
  }
}

function formatInstalls(installs) {
  if (!installs) return "N/A";
  return installs.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function getPlayStoreLink(appId) {
  return `https://play.google.com/store/apps/details?id=${appId}`;
}

async function fetchWithRetry(appId, retries = 3, delayMs = 1000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      return await gplay.app({ appId, lang: "en", country: "us" });
    } catch (err) {
      if (attempt < retries) {
        process.stdout.write(`    ↻ Retry ${attempt}/${retries - 1} for ${appId}...`);
        await delay(delayMs * attempt);
      } else {
        throw err;
      }
    }
  }
}


async function scrapeAll() {
  const allResults = [];

  for (let i = 0; i < SEARCH_TERMS.length; i++) {
    const term = SEARCH_TERMS[i];
    console.log(`\n[${i + 1}/${SEARCH_TERMS.length}] Searching: "${term}"`);

    try {
      const results = await gplay.search({
        term,
        num: RESULTS_PER_TERM,
        lang: "en",
        country: "us",
        fullDetail: false,
      });

      let count = 0;
      for (const app of results) {
        try {
          const detail = await fetchWithRetry(app.appId);
          allResults.push({
            searchTerm: term,
            appName: detail.title || "N/A",
            link: getPlayStoreLink(detail.appId),
            category: detail.genre || "N/A",
            lastUpdate: formatDate(detail.updated),
            rating: detail.score ? detail.score.toFixed(1) : "N/A",
            downloads: formatInstalls(detail.minInstalls),
            free: detail.free ? "Yes" : "No",
          });
          count++;
        } catch (detailErr) {
          // fallback to search-level data if detail call fails
          allResults.push({
            searchTerm: term,
            appName: app.title || "N/A",
            link: getPlayStoreLink(app.appId),
            category: app.genre || "N/A",
            lastUpdate: formatDate(app.updated),
            rating: app.score ? app.score.toFixed(1) : "N/A",
            downloads: formatInstalls(app.minInstalls),
            free: app.free ? "Yes" : "No",
          });
          count++;
        }
        await delay(300); // small delay between each app detail call
      }

      console.log(`  ✓ ${count} results found`);
    } catch (err) {
      console.error(`  ✗ Error for "${term}": ${err.message}`);
    }

    if (i < SEARCH_TERMS.length - 1) {
      process.stdout.write(`  Waiting ${DELAY_MS / 1000}s...`);
      await delay(DELAY_MS);
      console.log(" done");
    }
  }

  return allResults;
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
async function exportToExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "PlayStore Scraper";
  wb.created = new Date();

  const ws = wb.addWorksheet("Results");

  // Column definitions
  ws.columns = [
    { header: "Search Term",  key: "searchTerm", width: 30 },
    { header: "App Name",     key: "appName",    width: 35 },
    { header: "Link",         key: "link",       width: 60 },
    { header: "Category",     key: "category",   width: 25 },
    { header: "Last Update",  key: "lastUpdate", width: 20 },
    { header: "Rating",       key: "rating",     width: 10 },
    { header: "Downloads",    key: "downloads",  width: 18 },
    { header: "Free",         key: "free",       width: 8  },
  ];

  // Header style
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 11 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1A73E8" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
  });
  headerRow.height = 28;

  // Data rows
  let currentTerm = null;
  let rowColor = "FFFAFAFA";

  data.forEach((item, idx) => {
    if (item.searchTerm !== currentTerm) {
      currentTerm = item.searchTerm;
      rowColor = rowColor === "FFFAFAFA" ? "FFE8F0FE" : "FFFAFAFA";
    }

    const row = ws.addRow(item);
    row.height = 20;

    row.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowColor } };
      cell.font = { name: "Arial", size: 10 };
      cell.alignment = { vertical: "middle", wrapText: false };
      cell.border = {
        bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });

    // Make link a hyperlink
    const linkCell = row.getCell("link");
    linkCell.value = { text: item.link, hyperlink: item.link };
    linkCell.font = { name: "Arial", size: 10, color: { argb: "FF1A73E8" }, underline: true };

    // Color-code rating
    const ratingCell = row.getCell("rating");
    if (item.rating !== "N/A") {
      const r = parseFloat(item.rating);
      if (r >= 4.0) ratingCell.font = { ...ratingCell.font, color: { argb: "FF2E7D32" }, bold: true };
      else if (r >= 3.0) ratingCell.font = { ...ratingCell.font, color: { argb: "FFF57F17" } };
      else ratingCell.font = { ...ratingCell.font, color: { argb: "FFC62828" } };
    }

    // Color-code Free column
    const freeCell = row.getCell("free");
    freeCell.alignment = { horizontal: "center", vertical: "middle" };
    if (item.free === "Yes") freeCell.font = { ...freeCell.font, color: { argb: "FF2E7D32" }, bold: true };
    else freeCell.font = { ...freeCell.font, color: { argb: "FFC62828" } };
  });

  // Freeze header row
  ws.views = [{ state: "frozen", ySplit: 1 }];

  // Auto filter
  ws.autoFilter = { from: "A1", to: "H1" };

  // Summary sheet
  const summaryWs = wb.addWorksheet("Summary");
  summaryWs.columns = [
    { header: "Search Term", key: "term",  width: 35 },
    { header: "Apps Found",  key: "count", width: 15 },
  ];

  const summaryHeader = summaryWs.getRow(1);
  summaryHeader.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 11 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1A73E8" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
  summaryHeader.height = 28;

  const termCounts = {};
  data.forEach((r) => {
    termCounts[r.searchTerm] = (termCounts[r.searchTerm] || 0) + 1;
  });

  Object.entries(termCounts).forEach(([term, count]) => {
    const row = summaryWs.addRow({ term, count });
    row.height = 20;
    row.eachCell((c) => {
      c.font = { name: "Arial", size: 10 };
      c.alignment = { vertical: "middle" };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFAFAFA" } };
    });
  });

  // Total row
  const totalRow = summaryWs.addRow({ term: "TOTAL (unique apps)", count: data.length });
  totalRow.eachCell((c) => {
    c.font = { name: "Arial", size: 10, bold: true };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F0FE" } };
    c.alignment = { vertical: "middle" };
  });

  await wb.xlsx.writeFile(OUTPUT_FILE);
  console.log(`\n✅ Saved to: ${path.resolve(OUTPUT_FILE)}`);
}

// ─── Main ─────────────────────────────────────────────────────────────────────
(async () => {
  console.log("🚀 Starting Play Store scraper...");
  console.log(`   Terms: ${SEARCH_TERMS.length} | Results per term: ${RESULTS_PER_TERM}`);

  const data = await scrapeAll();

  if (data.length === 0) {
    console.error("No data collected. Exiting.");
    process.exit(1);
  }

  console.log(`\n📊 Total apps collected: ${data.length}`);
  console.log("📝 Writing Excel file...");
  await exportToExcel(data);
})();
