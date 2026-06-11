import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL("../outputs/profit-hunter-money-leak-finder/", import.meta.url);
const workbook = Workbook.create();

const theme = {
  ink: "#17231F",
  muted: "#60594F",
  paper: "#FBFAF5",
  warm: "#F2ECDF",
  line: "#DCD4C7",
  green: "#193D34",
  green2: "#2F755F",
  rust: "#C96243",
  gold: "#D6A63D",
  blue: "#4E719D",
  rose: "#F1D6CC",
  mint: "#E6F0DE",
  sky: "#DDEAF1",
  cream: "#FFF7E8",
  white: "#FFFFFF",
};

const currency = '"$"#,##0';
const percent = "0%";

const widths = {
  A: 24,
  B: 170,
  C: 150,
  D: 150,
  E: 150,
  F: 150,
  G: 150,
  H: 170,
  I: 180,
  J: 220,
  K: 180,
  L: 160,
};

const addSheet = (name) => {
  const sheet = workbook.worksheets.add(name);
  Object.entries(widths).forEach(([col, px]) => {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
  });
  sheet.getRange("A1:L80").format.font.name = "Aptos";
  sheet.getRange("A1:L80").format.font.size = 11;
  sheet.getRange("A1:L80").format.wrapText = true;
  sheet.getRange("A1:L80").format.font.color = theme.ink;
  return sheet;
};

const merge = (sheet, ranges) => ranges.forEach((range) => sheet.getRange(range).merge());

const titleBlock = (sheet, title, subtitle, step, range = "B2:K4") => {
  merge(sheet, ["B2:K2", "B3:K3", "B4:K4"]);
  sheet.getRange("B2:K2").values = [[title, "", "", "", "", "", "", "", "", ""]];
  sheet.getRange("B3:K3").values = [[subtitle, "", "", "", "", "", "", "", "", ""]];
  sheet.getRange("B4:K4").values = [[step, "", "", "", "", "", "", "", "", ""]];
  sheet.getRange(range).format.fill = theme.green;
  sheet.getRange("B2:K2").format.font = { color: theme.white, bold: true, size: 24 };
  sheet.getRange("B3:K3").format.font = { color: "#D8E7DE", size: 12 };
  sheet.getRange("B4:K4").format.font = { color: "#F6E2D8", bold: true, size: 11 };
};

const section = (sheet, range, label, fill = theme.warm) => {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[label]];
  sheet.getRange(range).format.fill = fill;
  sheet.getRange(range).format.font = { bold: true, color: theme.ink, size: 13 };
};

const header = (sheet, range) => {
  sheet.getRange(range).format.fill = theme.green;
  sheet.getRange(range).format.font = { color: theme.white, bold: true };
  sheet.getRange(range).format.horizontalAlignment = "center";
};

const card = (sheet, range, fill = theme.paper) => {
  sheet.getRange(range).format.fill = fill;
  sheet.getRange(range).format.borders = { preset: "outside", style: "thin", color: theme.line };
};

const moneyRows = [
  ["Monthly take-home pay", 4850, "What actually hits your bank account after payroll deductions."],
  ["Fixed bills", 2940, "Rent/mortgage, utilities, insurance, phone, internet, required payments."],
  ["Debt minimums", 520, "Minimum payments only. Extra payoff is handled separately."],
  ["Subscriptions", 218, "Recurring charges that quietly drain cash flow."],
  ["Savings transfer target", 400, "Automatic emergency fund or short-term savings transfer."],
  ["Groceries / gas / essentials", 650, "Normal living costs before fun spending."],
  ["Planned life money", 300, "Restaurants, fun, clothes, small wants, and personal spending."],
  ["Irregular expense buffer", 250, "Car repairs, gifts, annual fees, school costs, travel, etc."],
  ["Current emergency savings", 700, "Current cash set aside for emergencies."],
  ["Starter emergency target", 1000, "First milestone before getting aggressive elsewhere."],
];

const billsRows = [
  [1, "Rent / mortgage", 1850, "Housing", "Yes", "Paycheck 1", "Largest fixed bill"],
  [5, "Internet", 75, "Utilities", "Yes", "Paycheck 1", ""],
  [8, "Car payment", 420, "Transportation", "Yes", "Paycheck 1", ""],
  [12, "Phone", 95, "Utilities", "Yes", "Paycheck 1", ""],
  [15, "Utilities", 190, "Utilities", "No", "Paycheck 2", "Average monthly amount"],
  [18, "Insurance", 160, "Insurance", "Yes", "Paycheck 2", ""],
  [22, "Credit card minimum", 210, "Debt", "Yes", "Paycheck 2", ""],
  [25, "Student loan", 310, "Debt", "Yes", "Paycheck 2", ""],
];

const debtRows = [
  ["Credit Card 1", 4200, 24.99, 140, 125, "High interest"],
  ["Credit Card 2", 1800, 19.99, 70, 50, "Smallest balance"],
  ["Student Loan", 14500, 6.5, 310, 0, "Steady minimum"],
  ["Car Loan", 9300, 7.2, 420, 0, "Fixed payment"],
];

const subsRows = [
  ["Streaming bundle", 39, "Review", "Overlaps with another subscription"],
  ["Music", 11, "Keep", "Used daily"],
  ["Fitness app", 19, "Cut", "Not used this month"],
  ["Cloud storage", 10, "Keep", "Needed"],
  ["Meal plan app", 29, "Review", "Could pause"],
  ["News app", 7, "Cut", "Rarely used"],
  ["Design app", 35, "Review", "Check if still needed"],
  ["Membership", 68, "Review", "Annual value unclear"],
];

const start = addSheet("Start Here");
titleBlock(
  start,
  "Money Leak Finder",
  "A guided paycheck audit for people who earn enough but still feel squeezed.",
  "Start here: follow the tabs from left to right."
);
section(start, "B6:K6", "How this workbook works");
start.getRange("B7:K11").values = [
  ["1. Start Here", "Read the flow and what each tab is for.", "", "", "", "", "", "", "", ""],
  ["2. Dashboard", "See the headline numbers after you enter your data.", "", "", "", "", "", "", "", ""],
  ["3. Paycheck Map", "Enter your monthly money numbers. These drive the dashboard.", "", "", "", "", "", "", "", ""],
  ["4. Bills & Timing", "List due dates so you can spot paycheck timing traps.", "", "", "", "", "", "", "", ""],
  ["5. Debt & Subscriptions", "Find quiet leaks and choose where to cut, pause, or focus extra payoff.", "", "", "", "", "", "", "", ""],
];
merge(start, ["C7:K7", "C8:K8", "C9:K9", "C10:K10", "C11:K11"]);
start.getRange("B7:B11").format.font.bold = true;
card(start, "B7:K11", theme.white);
section(start, "B13:K13", "What to do next");
start.getRange("B14:K17").values = [
  ["Step 1", "Go to the Paycheck Map tab and replace every example number with your own.", "", "", "", "", "", "", "", ""],
  ["Step 2", "Go to Bills & Timing and enter due dates for your actual bills.", "", "", "", "", "", "", "", ""],
  ["Step 3", "Use Debt & Subscriptions to find cuts and pick one payoff focus.", "", "", "", "", "", "", "", ""],
  ["Step 4", "Return to Dashboard and Action Plan for the next move.", "", "", "", "", "", "", "", ""],
];
merge(start, ["C14:K14", "C15:K15", "C16:K16", "C17:K17"]);
start.getRange("B14:B17").format.font.bold = true;
card(start, "B14:K17", theme.mint);
section(start, "B19:K19", "Important note");
start.getRange("B20:K21").values = [
  ["This workbook is educational. It is meant to show cash-flow patterns, not provide individualized financial, legal, or investment advice.", "", "", "", "", "", "", "", "", ""],
  ["Use it as a monthly money meeting tool. The magic is not the spreadsheet. The magic is seeing the numbers before the month starts making choices for you.", "", "", "", "", "", "", "", "", ""],
];
merge(start, ["B20:K20", "B21:K21"]);
card(start, "B20:K21", theme.cream);

const dash = addSheet("Dashboard");
titleBlock(
  dash,
  "Dashboard",
  "Your paycheck truth in one page. Review this after completing the input tabs.",
  "If one number feels uncomfortable, that is the leak telling you where to look."
);
section(dash, "B6:K6", "Money Snapshot");
dash.getRange("B7:C9").values = [["Take-home pay", "",], ["", ""], ["", ""]];
dash.getRange("D7:E9").values = [["Committed money", ""], ["", ""], ["", ""]];
dash.getRange("F7:G9").values = [["Money leak risk", ""], ["", ""], ["", ""]];
dash.getRange("H7:I9").values = [["True flex cash", ""], ["", ""], ["", ""]];
dash.getRange("J7:K9").values = [["Emergency gap", ""], ["", ""], ["", ""]];
merge(dash, ["B7:C7", "B8:C9", "D7:E7", "D8:E9", "F7:G7", "F8:G9", "H7:I7", "H8:I9", "J7:K7", "J8:K9"]);
dash.getRange("B8").formulas = [["='Paycheck Map'!C8"]];
dash.getRange("D8").formulas = [["='Paycheck Map'!G8"]];
dash.getRange("F8").formulas = [["='Paycheck Map'!G9"]];
dash.getRange("H8").formulas = [["='Paycheck Map'!G10"]];
dash.getRange("J8").formulas = [["='Paycheck Map'!G11"]];
card(dash, "B7:C9", theme.mint);
card(dash, "D7:E9", theme.sky);
card(dash, "F7:G9", theme.rose);
card(dash, "H7:I9", theme.cream);
card(dash, "J7:K9", theme.warm);
dash.getRange("B7:K7").format.font = { bold: true, color: theme.muted, size: 11 };
dash.getRange("B8:K9").format.font = { bold: true, color: theme.green, size: 22 };
dash.getRange("B8:K9").format.numberFormat = currency;
dash.getRange("F8:G9").format.font.color = theme.rust;

section(dash, "B12:E12", "Paycheck Allocation");
dash.getRange("B13:E18").values = [
  ["Category", "Amount", "% of pay", "Signal"],
  ["Bills", "", "", ""],
  ["Debt minimums", "", "", ""],
  ["Subscriptions", "", "", ""],
  ["Savings target", "", "", ""],
  ["Life + essentials", "", "", ""],
];
dash.getRange("C14:C18").formulas = [["='Paycheck Map'!C9"], ["='Paycheck Map'!C10"], ["='Paycheck Map'!C11"], ["='Paycheck Map'!C12"], ["='Paycheck Map'!C13+'Paycheck Map'!C14+'Paycheck Map'!C15"]];
dash.getRange("D14:D18").formulas = [["=C14/$B$8"], ["=C15/$B$8"], ["=C16/$B$8"], ["=C17/$B$8"], ["=C18/$B$8"]];
dash.getRange("E14:E18").formulas = [
  ["=IF(D14>0.55,\"Heavy\",IF(D14>0.4,\"Watch\",\"OK\"))"],
  ["=IF(D15>0.15,\"Heavy\",IF(D15>0.08,\"Watch\",\"OK\"))"],
  ["=IF(D16>0.05,\"Leak check\",IF(D16>0.03,\"Watch\",\"OK\"))"],
  ["=IF(D17<0.05,\"Too low\",\"OK\")"],
  ["=IF(D18>0.35,\"Tight\",IF(D18>0.25,\"Watch\",\"OK\"))"],
];
header(dash, "B13:E13");
dash.getRange("C14:C18").format.numberFormat = currency;
dash.getRange("D14:D18").format.numberFormat = percent;
card(dash, "B13:E18", theme.white);
dash.getRange("D14:D18").conditionalFormats.add("dataBar", {
  color: theme.green2,
  gradient: true,
});

section(dash, "G12:K12", "What the numbers mean");
dash.getRange("G13:K18").values = [
  ["Money leak risk", "", "", "", ""],
  ["If this is above $0, your paycheck is overbooked before the month starts.", "", "", "", ""],
  ["True flex cash", "", "", "", ""],
  ["This is the amount you can spend without pretending bills or savings are optional.", "", "", "", ""],
  ["First move", "", "", "", ""],
  ["", "", "", "", ""],
];
dash.getRange("G18").formulas = [["='Paycheck Map'!G13"]];
merge(dash, ["G13:K13", "G14:K14", "G15:K15", "G16:K16", "G17:K17", "G18:K18"]);
dash.getRange("G13:G17").format.font.bold = true;
dash.getRange("G18:K18").format.font = { bold: true, color: theme.green, size: 14 };
card(dash, "G13:K18", theme.cream);

section(dash, "B21:K21", "Mini money meeting");
dash.getRange("B22:K25").values = [
  ["1", "What number surprised me most?", "", "", "", "", "", "", "", ""],
  ["2", "What bill timing problem can I fix this month?", "", "", "", "", "", "", "", ""],
  ["3", "What subscription can I cut, pause, or justify?", "", "", "", "", "", "", "", ""],
  ["4", "What one move will make next payday calmer?", "", "", "", "", "", "", "", ""],
];
merge(dash, ["C22:K22", "C23:K23", "C24:K24", "C25:K25"]);
dash.getRange("B22:B25").format.font.bold = true;
card(dash, "B22:K25", theme.mint);

const map = addSheet("Paycheck Map");
titleBlock(
  map,
  "Paycheck Map",
  "Enter the monthly numbers that drive the dashboard.",
  "Purpose: separate committed money from actual flex cash."
);
section(map, "B6:E6", "Monthly inputs");
map.getRange("B7:D17").values = [["Line Item", "Monthly Amount", "What this means"], ...moneyRows];
header(map, "B7:D7");
map.getRange("C8:C17").format.numberFormat = currency;
card(map, "B7:D17", theme.white);
section(map, "F6:K6", "Calculated results");
map.getRange("F7:H13").values = [
  ["Result", "Amount", "Why it matters"],
  ["Total committed money", "", "Bills, debt, subscriptions, savings, essentials, life money, and irregular expenses."],
  ["Money leak / overspend risk", "", "If above $0, the paycheck is overbooked."],
  ["True flex cash", "", "The real leftover after the paycheck has jobs."],
  ["Emergency fund gap", "", "How far you are from the starter emergency target."],
  ["Committed % of pay", "", "How much of take-home pay is already spoken for."],
  ["Recommended first move", "", "Your next concrete action."],
];
map.getRange("G8:G12").formulas = [["=SUM(C9:C15)"], ["=MAX(0,G8-C8)"], ["=MAX(0,C8-G8)"], ["=MAX(0,C17-C16)"], ["=G8/C8"]];
map.getRange("G13").formulas = [["=IF(G9>0,\"Cut or reschedule $\"&TEXT(G9,\"#,##0\"),IF(G11>0,\"Move $\"&TEXT(MIN(G10,G11),\"#,##0\")&\" toward starter emergency fund\",\"Protect this rhythm\"))"]];
map.getRange("G8:G11").format.numberFormat = currency;
map.getRange("G12").format.numberFormat = percent;
header(map, "F7:H7");
card(map, "F7:H13", theme.cream);
section(map, "B19:K19", "Next step");
map.getRange("B20:K21").values = [
  ["After this tab, go to Bills & Timing. The dashboard can show a leak, but the bill due dates often explain why it keeps happening.", "", "", "", "", "", "", "", "", ""],
  ["If any number is a guess, keep it for now. This workbook is meant to create visibility quickly, not punish you for imperfect data.", "", "", "", "", "", "", "", "", ""],
];
merge(map, ["B20:K20", "B21:K21"]);
card(map, "B20:K21", theme.mint);

const bills = addSheet("Bills & Timing");
titleBlock(
  bills,
  "Bills & Timing",
  "List due dates so you can spot paycheck timing traps.",
  "Purpose: find the bills that create the before-payday squeeze."
);
section(bills, "B6:J6", "Bill calendar");
bills.getRange("B7:H25").values = [
  ["Due Day", "Bill", "Amount", "Category", "Autopay?", "Covered By", "Notes"],
  ...billsRows,
  ...Array.from({ length: 10 }, () => ["", "", "", "", "", "", ""]),
];
header(bills, "B7:H7");
bills.getRange("D8:D25").format.numberFormat = currency;
card(bills, "B7:H25", theme.white);
section(bills, "B28:J28", "What to look for");
bills.getRange("B29:J32").values = [
  ["If most large bills hit before Paycheck 1 clears, timing may be the real problem.", "", "", "", "", "", "", "", ""],
  ["If possible, move one or two due dates so each paycheck carries a fair share.", "", "", "", "", "", "", "", ""],
  ["If a bill is on autopay, make sure the paycheck that covers it lands before the withdrawal.", "", "", "", "", "", "", "", ""],
  ["Next tab: Debt & Subscriptions. That is where the quiet monthly leaks usually show up.", "", "", "", "", "", "", "", ""],
];
merge(bills, ["B29:J29", "B30:J30", "B31:J31", "B32:J32"]);
card(bills, "B29:J32", theme.cream);

const debt = addSheet("Debt & Subscriptions");
titleBlock(
  debt,
  "Debt & Subscriptions",
  "Find the quiet leaks and choose one focus area.",
  "Purpose: reduce pressure without starving the rest of the month."
);
section(debt, "B6:H6", "Debt pressure");
debt.getRange("B7:H16").values = [
  ["Debt", "Balance", "Rate", "Minimum", "Extra Target", "Monthly Total", "Focus Note"],
  ...debtRows.map((row) => [row[0], row[1], row[2] / 100, row[3], row[4], "", row[5]]),
  ...Array.from({ length: 5 }, () => ["", "", "", "", "", "", ""]),
];
debt.getRange("G8:G16").formulas = Array.from({ length: 9 }, (_, i) => [`=IF(B${i + 8}=\"\",\"\",E${i + 8}+F${i + 8})`]);
header(debt, "B7:H7");
debt.getRange("C8:C16").format.numberFormat = currency;
debt.getRange("D8:D16").format.numberFormat = "0.00%";
debt.getRange("E8:G16").format.numberFormat = currency;
card(debt, "B7:H16", theme.white);
section(debt, "B19:H19", "Subscription audit");
debt.getRange("B20:G33").values = [
  ["Subscription", "Monthly Cost", "Keep / Cut / Review", "Annual Cost", "Reason", "Action"],
  ...subsRows.map((row) => [row[0], row[1], row[2], "", row[3], row[2] === "Cut" ? "Cancel this week" : row[2] === "Review" ? "Decide by payday" : "Keep"]),
  ...Array.from({ length: 5 }, () => ["", "", "", "", "", ""]),
];
debt.getRange("E21:E33").formulas = Array.from({ length: 13 }, (_, i) => [`=IF(C${i + 21}=\"\",\"\",C${i + 21}*12)`]);
header(debt, "B20:G20");
debt.getRange("C21:C33").format.numberFormat = currency;
debt.getRange("E21:E33").format.numberFormat = currency;
card(debt, "B20:G33", theme.white);
debt.getRange("E21:E33").conditionalFormats.add("dataBar", {
  color: theme.rust,
  gradient: true,
});
section(debt, "J6:K6", "Leak summary", theme.rose);
debt.getRange("J7:K12").values = [
  ["Total debt minimums", ""],
  ["Extra payoff target", ""],
  ["Subscription total", ""],
  ["Cut list total", ""],
  ["Annual subscription cost", ""],
  ["Next move", ""],
];
debt.getRange("K7:K11").formulas = [["=SUM(E8:E16)"], ["=SUM(F8:F16)"], ["=SUM(C21:C33)"], ["=SUMIF(D21:D33,\"Cut\",C21:C33)"], ["=SUM(E21:E33)"]];
debt.getRange("K12").formulas = [["=IF(K10>0,\"Cancel $\"&TEXT(K10,\"#,##0\")&\"/mo of unused subscriptions\",IF(K8>0,\"Apply extra payoff target intentionally\",\"Keep tracking\"))"]];
debt.getRange("K7:K11").format.numberFormat = currency;
card(debt, "J7:K12", theme.cream);

const action = addSheet("Action Plan");
titleBlock(
  action,
  "Action Plan",
  "Turn the audit into a simple next-payday plan.",
  "Purpose: leave with one move, not a pile of guilt."
);
section(action, "B6:K6", "Your next payday plan");
action.getRange("B7:C14").values = [
  ["Question", "Answer"],
  ["What is the first money leak to fix?", ""],
  ["What bill timing issue should I solve?", ""],
  ["What subscription will I cut, pause, or justify?", ""],
  ["How much will I move to savings next payday?", ""],
  ["What debt gets extra money only after bills and savings are covered?", ""],
  ["What is my true flex cash boundary?", ""],
  ["When is my next 20-minute money meeting?", ""],
];
merge(action, ["C7:K7", "C8:K8", "C9:K9", "C10:K10", "C11:K11", "C12:K12", "C13:K13", "C14:K14"]);
header(action, "B7:K7");
card(action, "B7:K14", theme.white);
action.getRange("C8:K14").format.fill = theme.paper;
action.getRange("C8:K14").format.borders = { preset: "inside", style: "thin", color: theme.line };
section(action, "B17:K17", "Suggested first move");
action.getRange("B18:K19").values = [["", "", "", "", "", "", "", "", "", ""], ["Use this as your one action before the next paycheck. Small, clean moves beat ambitious plans you never open again.", "", "", "", "", "", "", "", "", ""]];
action.getRange("B18").formulas = [["='Paycheck Map'!G13"]];
merge(action, ["B18:K18", "B19:K19"]);
action.getRange("B18:K18").format.font = { bold: true, color: theme.green, size: 16 };
card(action, "B18:K19", theme.mint);
section(action, "B22:K22", "The 20-minute monthly rhythm");
action.getRange("B23:K27").values = [
  ["1", "Update the Paycheck Map numbers.", "", "", "", "", "", "", "", ""],
  ["2", "Check bills due before each paycheck.", "", "", "", "", "", "", "", ""],
  ["3", "Review subscriptions and debt pressure.", "", "", "", "", "", "", "", ""],
  ["4", "Pick one money leak to fix.", "", "", "", "", "", "", "", ""],
  ["5", "Return to the dashboard and write the next move.", "", "", "", "", "", "", "", ""],
];
merge(action, ["C23:K23", "C24:K24", "C25:K25", "C26:K26", "C27:K27"]);
action.getRange("B23:B27").format.font.bold = true;
card(action, "B23:K27", theme.cream);

for (const sheet of [start, dash, map, bills, debt, action]) {
  sheet.getRange("B2:K4").format.borders = { preset: "outside", style: "thin", color: theme.green };
  sheet.getRange("B6:K70").format.borders = { preset: "inside", style: "thin", color: "#EFE7DA" };
  sheet.getRange("B6:K70").format.borders = { preset: "outside", style: "thin", color: theme.line };
  sheet.getRange("A:A").format.columnWidthPx = 24;
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const [sheetName, range] of [
  ["Start Here", "B2:K22"],
  ["Dashboard", "B2:K25"],
  ["Paycheck Map", "B2:K21"],
  ["Bills & Timing", "B2:J32"],
  ["Debt & Subscriptions", "B2:K33"],
  ["Action Plan", "B2:K27"],
]) {
  await workbook.render({ sheetName, range, scale: 1 });
}

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(new URL("money-leak-finder.xlsx", outputDir));
