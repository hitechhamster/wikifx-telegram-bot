import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workspace = path.resolve("..");
const outputDir = path.join(workspace, "outputs", "telegram_india_mvp_20260819");
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
workbook.comments.setSelf({ displayName: "User" });

const navy = "#0B2545";
const blue = "#1769AA";
const teal = "#07847E";
const lightBlue = "#EAF3FA";
const lightTeal = "#E7F6F3";
const lightYellow = "#FFF4CC";
const lightGray = "#F3F6F8";
const midGray = "#D6DEE4";
const red = "#B42318";

const setup = workbook.worksheets.add("Setup");
setup.showGridLines = false;
setup.getRange("A1:H1").merge();
setup.getRange("A1").values = [["Telegram India Broker-Check MVP"]];
setup.getRange("A1:H1").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF", size: 18 },
  rowHeight: 34,
  verticalAlignment: "center",
};
setup.getRange("A3:B12").values = [
  ["Item", "MVP decision"],
  ["Primary market", "India"],
  ["Launch language", "English"],
  ["Pakistan", "No separate group; about 1% incidental coverage"],
  ["Indonesia", "Clone the structure after the India MVP passes"],
  ["Data source", "This workbook (manual refresh)"],
  ["Retention", "30 days for Telegram user ID and query history"],
  ["P0 functions", "/start, /check, inline query, duplicate candidates, admin stats/export"],
  ["Deferred", "WATCH, automated risk alerts, multilingual replies, live API"],
  ["North-star", "App download followed by a verified broker-detail deep visit"],
];
setup.getRange("A3:B3").format = {
  fill: blue,
  font: { bold: true, color: "#FFFFFF" },
};
setup.getRange("A4:A12").format = { fill: lightBlue, font: { bold: true, color: navy } };
setup.getRange("B4:B12").format = { wrapText: true };
setup.getRange("A14:H14").merge();
setup.getRange("A14").values = [["Measurement boundary"]];
setup.getRange("A14:H14").format = { fill: teal, font: { bold: true, color: "#FFFFFF" } };
setup.getRange("A15:H17").merge();
setup.getRange("A15").values = [[
  "The bot can directly record who queried which broker, match outcomes and result delivery. A normal Telegram URL button does not prove that the user opened the page, installed the app or completed an in-app broker detail visit. Those stages require a tracking redirect plus GA4/App events and an AppsFlyer OneLink configured for broker-level deep linking.",
]];
setup.getRange("A15:H17").format = {
  fill: lightYellow,
  font: { color: "#5D4200" },
  wrapText: true,
  verticalAlignment: "center",
};
setup.getRange("A19:H19").merge();
setup.getRange("A19").values = [["Public-test gate"]];
setup.getRange("A19:H19").format = { fill: red, font: { bold: true, color: "#FFFFFF" } };
setup.getRange("A20:H22").merge();
setup.getRange("A20").values = [[
  "Internal testing can begin immediately. Before a public launch, add a privacy notice, confirm brand/legal wording, validate each broker row against the live WikiFX page, and confirm whether the current OneLink can carry broker_id and return an installed user to the exact broker page.",
]];
setup.getRange("A20:H22").format = { fill: "#FEECEB", font: { color: red }, wrapText: true };
setup.getRange("A1:H22").format.borders = { preset: "outside", style: "thin", color: midGray };
setup.getRange("A:A").format.columnWidth = 24;
setup.getRange("B:B").format.columnWidth = 72;
setup.getRange("C:H").format.columnWidth = 12;
setup.freezePanes.freezeRows(3);

const brokers = workbook.worksheets.add("India Brokers");
brokers.showGridLines = false;
brokers.getRange("A1:N1").merge();
brokers.getRange("A1").values = [["India MVP broker source — editable inputs"]];
brokers.getRange("A1:N1").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF", size: 16 },
  rowHeight: 32,
  verticalAlignment: "center",
};
brokers.getRange("A2:N2").merge();
brokers.getRange("A2").values = [[
  "Yellow cells are manually maintained. Live scores, complaints and regulatory details can change; verify the source URL immediately before public use.",
]];
brokers.getRange("A2:N2").format = { fill: lightYellow, font: { color: "#5D4200" }, wrapText: true };
brokers.getRange("A4:N4").values = [[
  "broker_id",
  "broker_name",
  "aliases",
  "official_domains",
  "license_numbers",
  "regulation_status",
  "risk_summary",
  "complaint_count",
  "wikifx_url",
  "app_download_url",
  "market",
  "language",
  "source_checked_at",
  "active",
]];

const appLink = "https://wikifx2.onelink.me/ohOG/zeq80ve1";
const brokerRows = [
  [
    "0001390005",
    "Exness",
    "exness|exness broker",
    "exness.com",
    "",
    "Regulated in several non-Indian jurisdictions; not locally authorised for India. Verify the live WikiFX profile.",
    "India users should verify current local eligibility and regulatory protection before taking any action.",
    null,
    "https://www.wikifx.com/en/dealer/0001390005.html",
    appLink,
    "India",
    "English",
    new Date("2026-08-19T00:00:00+08:00"),
    true,
  ],
  [
    "0001461138",
    "XM",
    "xm|xm global|xm broker",
    "xm.com",
    "ASIC|CySEC|DFSA|FSC",
    "Multiple international regulatory records are shown on WikiFX; verify the exact entity serving the user.",
    "Regulatory protection depends on the contracted entity and jurisdiction, not the brand name alone.",
    null,
    "https://www.wikifx.com/en/dealer/0001461138.html",
    appLink,
    "India",
    "English",
    new Date("2026-08-19T00:00:00+08:00"),
    true,
  ],
  [
    "1941447945",
    "Octa",
    "octa|octafx|octa fx",
    "octafx.eu|octa.com",
    "CySEC 372/18",
    "CySEC-regulated entity is visible; the live profile also shows a high-potential-risk warning.",
    "WikiFX reported a high complaint volume on the checked profile. Confirm the latest count before publishing.",
    168,
    "https://www.wikifx.com/en/dealer/1941447945.html",
    appLink,
    "India",
    "English",
    new Date("2026-08-19T00:00:00+08:00"),
    true,
  ],
  [
    "9641842942",
    "IC Markets Global",
    "ic markets|ic markets global|icmarkets|ic",
    "icmarkets.com",
    "ASIC 335692|CySEC 362/18|FSA SD018",
    "International regulation shown for Australia, Cyprus and Seychelles; verify the entity serving India.",
    "Offshore and local protection differ by contracting entity. Review the exact licence and client entity.",
    null,
    "https://www.wikifx.com/en/dealer/9641842942.html",
    appLink,
    "India",
    "English",
    new Date("2026-08-19T00:00:00+08:00"),
    true,
  ],
  [
    "2277676718",
    "STARTRADER",
    "startrader|star trader",
    "startrader.com|startraderprime.com.au",
    "ASIC|FSCA|FSA",
    "International regulatory records are shown on WikiFX; verify the exact entity and current status.",
    "Do not treat a global licence as proof of Indian authorisation. Confirm the user-facing entity.",
    null,
    "https://www.wikifx.com/en/dealer/2277676718.html",
    appLink,
    "India",
    "English",
    new Date("2026-08-19T00:00:00+08:00"),
    true,
  ],
];
brokers.getRange("A5:N9").values = brokerRows;
brokers.getRange("A4:N4").format = {
  fill: blue,
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  rowHeight: 32,
  verticalAlignment: "center",
};
brokers.getRange("A5:N9").format = {
  fill: lightYellow,
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "inside", style: "thin", color: midGray },
};
brokers.getRange("M5:M9").format.numberFormat = "yyyy-mm-dd";
brokers.getRange("H5:H9").format.numberFormat = "0";
brokers.getRange("N5:N100").dataValidation = { rule: { type: "list", values: [true, false] } };
brokers.getRange("K5:K100").dataValidation = { rule: { type: "list", values: ["India", "Pakistan", "Indonesia"] } };
brokers.getRange("L5:L100").dataValidation = { rule: { type: "list", values: ["English", "Bahasa Indonesia", "Hindi", "Urdu"] } };
brokers.freezePanes.freezeRows(4);
brokers.freezePanes.freezeColumns(2);
const widths = [16, 23, 30, 26, 30, 48, 52, 16, 48, 42, 14, 18, 18, 12];
for (let i = 0; i < widths.length; i += 1) {
  brokers.getRangeByIndexes(0, i, 9, 1).format.columnWidth = widths[i];
}
brokers.tables.add("A4:N9", true, "IndiaBrokerTable").style = "TableStyleMedium2";
brokers.getRange("F5").addComment?.("Regulatory statements must be checked against the live source before public publication.");

const schema = workbook.worksheets.add("Event Schema");
schema.showGridLines = false;
schema.getRange("A1:F1").merge();
schema.getRange("A1").values = [["Telegram event and reporting contract"]];
schema.getRange("A1:F1").format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 32 };
schema.getRange("A3:F3").values = [["event_type", "when", "required fields", "directly measurable", "retention", "notes"]];
schema.getRange("A4:F12").values = [
  ["bot_start", "User runs /start", "user_id, username, chat_id, market", "Yes — bot database", "30 days", "Show privacy notice"],
  ["broker_query", "User submits name/domain/licence", "user_id, query_text, chat_id", "Yes — bot database", "30 days", "Core demand signal"],
  ["match_success", "One broker matched", "broker_id, match_type", "Yes — bot database", "30 days", "Name/domain/licence"],
  ["match_multiple", "More than one candidate", "candidate_count", "Yes — bot database", "30 days", "User selects candidate"],
  ["match_none", "No candidate", "query_text", "Yes — bot database", "30 days", "No human escalation in P0"],
  ["result_sent", "Summary card returned", "broker_id, source_tag", "Yes — bot database", "30 days", "Does not prove click"],
  ["detail_click", "User presses details CTA", "broker_id, click_id", "Not with a plain URL button", "30 days", "Needs redirect or 2-step callback"],
  ["web_deep_view", "Exact broker page loads", "broker_id, source_id, page_type", "Requires GA4/site event", "Per analytics policy", "Must be added to the website"],
  ["app_download", "Store/install attribution", "campaign, broker_id, source_id", "Requires AppsFlyer OneLink", "Per AppsFlyer policy", "Existing generic OneLink found; broker-level setup unverified"],
];
schema.getRange("A3:F3").format = { fill: teal, font: { bold: true, color: "#FFFFFF" }, wrapText: true };
schema.getRange("A4:F12").format = { wrapText: true, verticalAlignment: "top" };
schema.getRange("A3:F12").format.borders = { preset: "all", style: "thin", color: midGray };
schema.getRange("A:A").format.columnWidth = 22;
schema.getRange("B:B").format.columnWidth = 30;
schema.getRange("C:C").format.columnWidth = 40;
schema.getRange("D:D").format.columnWidth = 32;
schema.getRange("E:E").format.columnWidth = 20;
schema.getRange("F:F").format.columnWidth = 48;
schema.freezePanes.freezeRows(3);

const dashboard = workbook.worksheets.add("Dashboard Spec");
dashboard.showGridLines = false;
dashboard.getRange("A1:H1").merge();
dashboard.getRange("A1").values = [["India MVP dashboard — acceptance metrics"]];
dashboard.getRange("A1:H1").format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 32 };
dashboard.getRange("A3:D3").values = [["Stage", "Metric", "Source", "Status before integration"]];
dashboard.getRange("A4:D11").values = [
  ["Telegram", "Unique users who queried", "Bot SQLite", "Ready in scaffold"],
  ["Telegram", "Queries per broker", "Bot SQLite", "Ready in scaffold"],
  ["Telegram", "Match success / multiple / none", "Bot SQLite", "Ready in scaffold"],
  ["Telegram", "Admin CSV export", "Bot /export command", "Ready in scaffold"],
  ["Web", "Exact broker detail visits", "GA4 custom event", "Needs site validation"],
  ["App", "App-store click / install", "AppsFlyer OneLink", "Generic OneLink exists"],
  ["App", "Exact broker detail deep visit", "AppsFlyer + app event", "Unverified; product/data confirmation required"],
  ["North-star", "Download followed by verified broker deep visit", "Joined attribution", "Not measurable yet"],
];
dashboard.getRange("A3:D3").format = { fill: blue, font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("A4:D11").format = { wrapText: true, verticalAlignment: "top" };
dashboard.getRange("A3:D11").format.borders = { preset: "all", style: "thin", color: midGray };
dashboard.getRange("F3:H3").merge();
dashboard.getRange("F3").values = [["One-week internal pilot targets"]];
dashboard.getRange("F3:H3").format = { fill: teal, font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("F4:G10").values = [
  ["Test accounts", 5],
  ["Valid query scenarios", 15],
  ["Name-match tests", 5],
  ["Domain-match tests", 5],
  ["Licence-match tests", 3],
  ["Multi/no-match tests", 2],
  ["Wrong-broker matches", 0],
];
dashboard.getRange("F4:F10").format = { fill: lightTeal, font: { bold: true, color: navy } };
dashboard.getRange("G4:G10").format.numberFormat = "0";
dashboard.getRange("F4:G10").format.borders = { preset: "all", style: "thin", color: midGray };
dashboard.getRange("A:A").format.columnWidth = 18;
dashboard.getRange("B:B").format.columnWidth = 38;
dashboard.getRange("C:C").format.columnWidth = 28;
dashboard.getRange("D:D").format.columnWidth = 38;
dashboard.getRange("E:E").format.columnWidth = 4;
dashboard.getRange("F:F").format.columnWidth = 28;
dashboard.getRange("G:G").format.columnWidth = 14;
dashboard.getRange("H:H").format.columnWidth = 14;

const previewNames = [
  ["Setup", "setup.png", "A1:H22"],
  ["India Brokers", "brokers.png", "A1:N9"],
  ["Event Schema", "schema.png", "A1:F12"],
  ["Dashboard Spec", "dashboard.png", "A1:H11"],
];
for (const [sheetName, fileName, range] of previewNames) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, fileName), new Uint8Array(await preview.arrayBuffer()));
}

const workbookFile = await SpreadsheetFile.exportXlsx(workbook);
await workbookFile.save(path.join(outputDir, "Telegram_India_Broker_MVP.xlsx"));

const check = await workbook.inspect({
  kind: "table",
  range: "India Brokers!A4:N9",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 14,
  maxChars: 9000,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

