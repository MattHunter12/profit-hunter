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
  B: 180,
  C: 130,
  D: 130,
  E: 130,
  F: 130,
  G: 160,
  H: 230,
  I: 170,
  J: 210,
  K: 180,
  L: 160,
};

const addSheet = (name) => {
  const sheet = workbook.worksheets.add(name);
  Object.entries(widths).forEach(([col, px]) => {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
  });
  sheet.getRange("A1:L90").format.font.name = "Aptos";
  sheet.getRange("A1:L90").format.font.size = 11;
  sheet.getRange("A1:L90").format.wrapText = true;
  sheet.getRange("A1:L90").format.font.color = theme.ink;
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

const planRows = [
  ["Income", "Income", 4850, "All paychecks and cash received this month."],
  ["Bills", "Expense", 2940, "Rent/mortgage, utilities, insurance, phone, internet, and required bills."],
  ["Debt", "Expense", 520, "Minimum debt payments. Extra payoff is planned on the debt tab."],
  ["Subscriptions", "Expense", 218, "Recurring charges that quietly drain cash flow."],
  ["Savings Transfers", "Expense", 400, "Money moved to emergency savings, sinking funds, or investments."],
  ["Groceries", "Expense", 500, "Food at home and normal household basics."],
  ["Transportation", "Expense", 150, "Gas, parking, rideshare, tolls, and transit."],
  ["Pets", "Expense", 90, "Dog food, pet supplies, grooming, vet copays, and similar costs."],
  ["Food & Dining", "Expense", 250, "Restaurants, coffee, airport sandwiches, and convenience meals."],
  ["Shopping / Life", "Expense", 300, "Clothes, personal care, small wants, and everyday life spending."],
  ["Travel / Events", "Expense", 125, "Trips, airport costs, activities, concerts, and events."],
  ["Irregular Expenses", "Expense", 250, "Repairs, gifts, annual fees, school costs, and surprise spending."],
  ["Current Emergency Savings", "Info", "", "Enter the cash you currently have set aside for emergencies."],
  ["Starter Emergency Target", "Info", "", "A simple first milestone before getting aggressive elsewhere."],
];

const logRows = [
  ["2026-06-01", "Paycheck 1", "Income", 2450, "Income", "", "Example income entry"],
  ["2026-06-01", "Rent / mortgage", "Bills", 1850, "Expense", "Need", "Example fixed bill"],
  ["2026-06-03", "Grocery run", "Groceries", 112, "Expense", "Need", ""],
  ["2026-06-05", "Dog food", "Pets", 48, "Expense", "Need", "This is how normal real-life spending gets captured"],
  ["2026-06-08", "Airport sandwich", "Food & Dining", 18, "Expense", "Want", "Small leaks become visible here"],
  ["2026-06-10", "Streaming bundle", "Subscriptions", 39, "Expense", "Want", ""],
  ["2026-06-15", "Paycheck 2", "Income", 2400, "Income", "", "Example income entry"],
];

const billsRows = [
  [1, "Rent / mortgage", 1850, "Bills", "Yes", "Paycheck 1", "Largest fixed bill"],
  [5, "Internet", 75, "Bills", "Yes", "Paycheck 1", ""],
  [8, "Car payment", 420, "Debt", "Yes", "Paycheck 1", ""],
  [12, "Phone", 95, "Bills", "Yes", "Paycheck 1", ""],
  [15, "Utilities", 190, "Bills", "No", "Paycheck 2", "Average monthly amount"],
  [18, "Insurance", 160, "Bills", "Yes", "Paycheck 2", ""],
  [22, "Credit card minimum", 210, "Debt", "Yes", "Paycheck 2", ""],
  [25, "Student loan", 310, "Debt", "Yes", "Paycheck 2", ""],
];

const debtRows = [
  ["Credit Card 1", 4200, 24.99, 140, 125, 22, "High interest"],
  ["Credit Card 2", 1800, 19.99, 70, 50, 8, "Smallest balance"],
  ["Student Loan", 14500, 6.5, 310, 0, 25, "Steady minimum"],
  ["Car Loan", 9300, 7.2, 420, 0, 8, "Fixed payment"],
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

const categoryGuide = [
  ["Income", "Paychecks, side income, reimbursements, and cash received.", "Income"],
  ["Bills", "Rent, utilities, insurance, phone, internet, and required bills.", "Expense"],
  ["Debt", "Debt minimums and extra payoff.", "Expense"],
  ["Subscriptions", "Recurring apps, memberships, software, and entertainment.", "Expense"],
  ["Savings Transfers", "Emergency fund, sinking funds, investing transfers.", "Expense"],
  ["Groceries", "Food at home and household basics.", "Expense"],
  ["Transportation", "Gas, parking, rideshare, tolls, transit.", "Expense"],
  ["Pets", "Pet food, supplies, grooming, vet copays.", "Expense"],
  ["Food & Dining", "Restaurants, coffee, airport meals, takeout.", "Expense"],
  ["Shopping / Life", "Clothes, personal care, gifts, everyday wants.", "Expense"],
  ["Travel / Events", "Trips, airports, concerts, sports, activities.", "Expense"],
  ["Irregular Expenses", "Repairs, annual fees, surprise costs.", "Expense"],
];

const start = addSheet("Start Here");
titleBlock(
  start,
  "Money Leak Finder",
  "A guided paycheck audit for people who earn enough but still feel squeezed.",
  "Start here: follow the tabs from left to right."
);
section(start, "B6:K6", "How this workbook works");
start.getRange("B7:K12").values = [
  ["1. Start Here", "Read the flow and what each tab is for.", "", "", "", "", "", "", "", ""],
  ["2. Dashboard", "Review the headline numbers after you enter your plan and actual spending.", "", "", "", "", "", "", "", ""],
  ["3. Paycheck Map", "Set the budget, then compare it to actuals pulled from the Daily Log.", "", "", "", "", "", "", "", ""],
  ["4. Daily Log", "Enter real income and spending as life happens. This feeds the actuals.", "", "", "", "", "", "", "", ""],
  ["5. Bills & Timing", "Plan due dates so you can spot paycheck timing traps.", "", "", "", "", "", "", "", ""],
  ["6. Debt & Subscriptions", "Enter debt details clearly and choose what to cut, pause, or attack.", "", "", "", "", "", "", "", ""],
];
merge(start, ["C7:K7", "C8:K8", "C9:K9", "C10:K10", "C11:K11", "C12:K12"]);
start.getRange("B7:B12").format.font.bold = true;
card(start, "B7:K12", theme.white);
section(start, "B14:K14", "What to do next");
start.getRange("B15:K19").values = [
  ["Step 1", "Set the budget on Paycheck Map. Keep it rough if you need speed.", "", "", "", "", "", "", "", ""],
  ["Step 2", "Use Daily Log for each paycheck, bill payment, grocery run, pet cost, meal, or impulse buy.", "", "", "", "", "", "", "", ""],
  ["Step 3", "Use Bills & Timing to see whether the problem is the amount, the timing, or both.", "", "", "", "", "", "", "", ""],
  ["Step 4", "Use Debt & Subscriptions to clarify balances, minimums, rates, and easy cuts.", "", "", "", "", "", "", "", ""],
  ["Step 5", "Return to Dashboard and Action Plan for the next concrete move.", "", "", "", "", "", "", "", ""],
];
merge(start, ["C15:K15", "C16:K16", "C17:K17", "C18:K18", "C19:K19"]);
start.getRange("B15:B19").format.font.bold = true;
card(start, "B15:K19", theme.mint);
section(start, "B22:K22", "Important note");
start.getRange("B23:K24").values = [
  ["This workbook is educational. It is meant to show cash-flow patterns, not provide individualized financial, legal, or investment advice.", "", "", "", "", "", "", "", "", ""],
  ["The daily log is where the trust is built: if spending gets entered there, the dashboard tells the truth.", "", "", "", "", "", "", "", "", ""],
];
merge(start, ["B23:K23", "B24:K24"]);
card(start, "B23:K24", theme.cream);

const dash = addSheet("Dashboard");
titleBlock(
  dash,
  "Dashboard",
  "Your plan vs actual paycheck truth in one page.",
  "Review this after updating the Paycheck Map and Daily Log."
);
section(dash, "B6:K6", "Money Snapshot");
dash.getRange("B7:C9").values = [["Actual income", ""], ["", ""], ["", ""]];
dash.getRange("D7:E9").values = [["Actual spending", ""], ["", ""], ["", ""]];
dash.getRange("F7:G9").values = [["True flex cash", ""], ["", ""], ["", ""]];
dash.getRange("H7:I9").values = [["Money leak risk", ""], ["", ""], ["", ""]];
dash.getRange("J7:K9").values = [["First move", ""], ["", ""], ["", ""]];
merge(dash, ["B7:C7", "B8:C9", "D7:E7", "D8:E9", "F7:G7", "F8:G9", "H7:I7", "H8:I9", "J7:K7", "J8:K9"]);
dash.getRange("B8").formulas = [["='Paycheck Map'!K8"]];
dash.getRange("D8").formulas = [["='Paycheck Map'!K9"]];
dash.getRange("F8").formulas = [["='Paycheck Map'!K10"]];
dash.getRange("H8").formulas = [["='Paycheck Map'!K11"]];
dash.getRange("J8").formulas = [["='Paycheck Map'!K14"]];
card(dash, "B7:C9", theme.mint);
card(dash, "D7:E9", theme.sky);
card(dash, "F7:G9", theme.cream);
card(dash, "H7:I9", theme.rose);
card(dash, "J7:K9", theme.warm);
dash.getRange("B7:K7").format.font = { bold: true, color: theme.muted, size: 11 };
dash.getRange("B8:I9").format.font = { bold: true, color: theme.green, size: 22 };
dash.getRange("B8:I9").format.numberFormat = currency;
dash.getRange("H8:I9").format.font.color = theme.rust;
dash.getRange("J8:K9").format.font = { bold: true, color: theme.green, size: 13 };

section(dash, "B12:F12", "Budget vs Actual");
dash.getRange("B13:F24").values = [
  ["Category", "Budget", "Actual", "Difference", "Signal"],
  ...Array.from({ length: 11 }, () => ["", "", "", "", ""]),
];
dash.getRange("B14:B24").formulas = Array.from({ length: 11 }, (_, i) => [`='Paycheck Map'!B${i + 9}`]);
dash.getRange("C14:C24").formulas = Array.from({ length: 11 }, (_, i) => [`='Paycheck Map'!D${i + 9}`]);
dash.getRange("D14:D24").formulas = Array.from({ length: 11 }, (_, i) => [`='Paycheck Map'!E${i + 9}`]);
dash.getRange("E14:E24").formulas = Array.from({ length: 11 }, (_, i) => [`='Paycheck Map'!F${i + 9}`]);
dash.getRange("F14:F24").formulas = Array.from({ length: 11 }, (_, i) => [`='Paycheck Map'!G${i + 9}`]);
header(dash, "B13:F13");
dash.getRange("C14:E24").format.numberFormat = currency;
card(dash, "B13:F24", theme.white);
dash.getRange("D14:D24").conditionalFormats.add("dataBar", {
  color: theme.green2,
  gradient: true,
});

section(dash, "H12:K12", "What the numbers mean");
dash.getRange("H13:K19").values = [
  ["Budget vs actual", "", "", ""],
  ["Positive difference means you are under budget. Negative means that category has become a leak.", "", "", ""],
  ["Daily Log", "", "", ""],
  ["The actual column is only as honest as the spending entered on the Daily Log tab.", "", "", ""],
  ["Most common leak", "", "", ""],
  ["", "", "", ""],
  ["", "", "", ""],
];
dash.getRange("H18").formulas = [["='Paycheck Map'!K13"]];
merge(dash, ["H13:K13", "H14:K14", "H15:K15", "H16:K16", "H17:K17", "H18:K19"]);
dash.getRange("H13:H17").format.font.bold = true;
dash.getRange("H18:K19").format.font = { bold: true, color: theme.green, size: 14 };
card(dash, "H13:K19", theme.cream);

section(dash, "B27:K27", "Mini money meeting");
dash.getRange("B28:K31").values = [
  ["1", "What category is over budget and why?", "", "", "", "", "", "", "", ""],
  ["2", "Was the issue a real need, a timing problem, or a convenience purchase?", "", "", "", "", "", "", "", ""],
  ["3", "What purchase pattern would I change before next payday?", "", "", "", "", "", "", "", ""],
  ["4", "What one move will make next payday calmer?", "", "", "", "", "", "", "", ""],
];
merge(dash, ["C28:K28", "C29:K29", "C30:K30", "C31:K31"]);
dash.getRange("B28:B31").format.font.bold = true;
card(dash, "B28:K31", theme.mint);

const map = addSheet("Paycheck Map");
titleBlock(
  map,
  "Paycheck Map",
  "Set the budget, then compare it to actuals from the Daily Log.",
  "Purpose: see what you planned, what happened, and where the leak started."
);
section(map, "B6:H6", "Budget vs Actual");
map.getRange("B7:H21").values = [
  ["Category", "Type", "Budget", "Actual", "Difference", "Signal", "What this means"],
  ...planRows.map(([category, type, budget, description]) => [category, type, budget, "", "", "", description]),
];
header(map, "B7:H7");
map.getRange("D8:D19").format.fill = theme.paper;
map.getRange("D20:E21").format.fill = theme.paper;
map.getRange("D8:E21").format.numberFormat = currency;
map.getRange("E8").formulas = [["=SUMIF('Daily Log'!F:F,\"Income\",'Daily Log'!E:E)"]];
map.getRange("E9:E19").formulas = Array.from({ length: 11 }, (_, i) => [`=SUMIF('Daily Log'!D:D,B${i + 9},'Daily Log'!E:E)`]);
map.getRange("E20:E21").values = [[700], [1000]];
map.getRange("F8:F19").formulas = Array.from({ length: 12 }, (_, i) => {
  const row = i + 8;
  return [`=IF(C${row}=\"Income\",E${row}-D${row},D${row}-E${row})`];
});
map.getRange("G8:G19").formulas = Array.from({ length: 12 }, (_, i) => {
  const row = i + 8;
  return [`=IF(C${row}=\"Income\",IF(F${row}<0,\"Under plan\",\"On track\"),IF(F${row}<0,\"Over budget\",\"OK\"))`];
});
map.getRange("F20:F21").formulas = [[""], ["=MAX(0,E21-E20)"]];
map.getRange("G20").values = [["Manual input"]];
map.getRange("G21").formulas = [["=IF(F21>0,\"Gap remains\",\"Funded\")"]];
map.getRange("D8:F21").format.numberFormat = currency;
card(map, "B7:H21", theme.white);
section(map, "B24:H24", "How actuals work");
map.getRange("B25:H27").values = [
  ["Actuals for Income and Expenses pull from the Daily Log. If a real-life transaction is missing there, it is missing from the dashboard.", "", "", "", "", "", ""],
  ["Use categories consistently. Dog food goes to Pets. Airport sandwich goes to Food & Dining. Rent goes to Bills. Credit card payment goes to Debt.", "", "", "", "", "", ""],
  ["Emergency savings is intentionally manual because it is a balance, not a spending transaction.", "", "", "", "", "", ""],
];
merge(map, ["B25:H25", "B26:H26", "B27:H27"]);
card(map, "B25:H27", theme.mint);
section(map, "J6:K6", "Calculated results");
map.getRange("J7:K14").values = [
  ["Result", "Amount / Signal"],
  ["Total income", ""],
  ["Total expenses", ""],
  ["True flex cash", ""],
  ["Money leak risk", ""],
  ["Emergency gap", ""],
  ["Most over-budget category", ""],
  ["Recommended first move", ""],
];
map.getRange("K8:K12").formulas = [
  ["=E8"],
  ["=SUM(E9:E19)"],
  ["=E8-SUM(E9:E19)"],
  ["=MAX(0,SUM(E9:E19)-E8)"],
  ["=F21"],
];
map.getRange("K13").formulas = [["=IF(MIN(F9:F19)<0,INDEX(B9:B19,MATCH(MIN(F9:F19),F9:F19,0)),\"No over-budget category yet\")"]];
map.getRange("K14").formulas = [["=IF(K11>0,\"Cut or reschedule $\"&TEXT(K11,\"#,##0\"),IF(K12>0,\"Move $\"&TEXT(MIN(K10,K12),\"#,##0\")&\" toward starter emergency fund\",\"Protect this rhythm\"))"]];
header(map, "J7:K7");
map.getRange("K8:K12").format.numberFormat = currency;
card(map, "J7:K14", theme.cream);

const log = addSheet("Daily Log");
titleBlock(
  log,
  "Daily Log",
  "Enter real income and spending as life happens.",
  "Purpose: feed actuals without making you update the same number twice."
);
section(log, "B6:H6", "Transaction log");
log.getRange("B7:H45").values = [
  ["Date", "Description", "Category", "Amount", "Type", "Need / Want", "Notes"],
  ...logRows,
  ...Array.from({ length: 31 }, () => ["", "", "", "", "", "", ""]),
];
header(log, "B7:H7");
log.getRange("B8:B45").format.numberFormat = "yyyy-mm-dd";
log.getRange("E8:E45").format.numberFormat = currency;
card(log, "B7:H45", theme.white);
section(log, "J6:K6", "Category guide");
log.getRange("J7:L19").values = [["Category", "Use for", "Type"], ...categoryGuide];
header(log, "J7:L7");
card(log, "J7:L19", theme.cream);
section(log, "B48:H48", "What counts here");
log.getRange("B49:H51").values = [
  ["Use this for groceries, dog food, coffee, airport meals, bills paid, paychecks received, savings transfers, debt payments, and subscriptions.", "", "", "", "", "", ""],
  ["The goal is not perfection. The goal is enough visibility to catch the leak before the next paycheck.", "", "", "", "", "", ""],
  ["If a category does not fit, use the closest category for now. The paid kit can go deeper later.", "", "", "", "", "", ""],
];
merge(log, ["B49:H49", "B50:H50", "B51:H51"]);
card(log, "B49:H51", theme.mint);

const bills = addSheet("Bills & Timing");
titleBlock(
  bills,
  "Bills & Timing",
  "List due dates so you can spot paycheck timing traps.",
  "Purpose: find the bills that create the before-payday squeeze."
);
section(bills, "B6:H6", "Bill calendar");
bills.getRange("B7:H25").values = [
  ["Due Day", "Bill", "Amount", "Log Category", "Autopay?", "Covered By", "Notes"],
  ...billsRows,
  ...Array.from({ length: 10 }, () => ["", "", "", "", "", "", ""]),
];
header(bills, "B7:H7");
bills.getRange("D8:D25").format.numberFormat = currency;
card(bills, "B7:H25", theme.white);
section(bills, "B28:K28", "What to look for");
bills.getRange("B29:K33").values = [
  ["This tab plans timing. Actual payments still belong in Daily Log using the Log Category shown above.", "", "", "", "", "", "", "", "", ""],
  ["If most large bills hit before Paycheck 1 clears, timing may be the real problem.", "", "", "", "", "", "", "", "", ""],
  ["If possible, move one or two due dates so each paycheck carries a fair share.", "", "", "", "", "", "", "", "", ""],
  ["If a bill is on autopay, make sure the paycheck that covers it lands before the withdrawal.", "", "", "", "", "", "", "", "", ""],
  ["Next tab: Debt & Subscriptions. That is where the quiet monthly leaks usually show up.", "", "", "", "", "", "", "", "", ""],
];
merge(bills, ["B29:K29", "B30:K30", "B31:K31", "B32:K32", "B33:K33"]);
card(bills, "B29:K33", theme.cream);

const debt = addSheet("Debt & Subscriptions");
titleBlock(
  debt,
  "Debt & Subscriptions",
  "Enter debt details clearly and choose one focus area.",
  "Purpose: reduce pressure without starving the rest of the month."
);
section(debt, "B6:I6", "Debt details");
debt.getRange("B7:I16").values = [
  ["Debt Name", "Current Balance", "Interest Rate / APR", "Minimum Payment", "Extra Target", "Due Day", "Monthly Total", "Notes"],
  ...debtRows.map((row) => [row[0], row[1], row[2] / 100, row[3], row[4], row[5], "", row[6]]),
  ...Array.from({ length: 5 }, () => ["", "", "", "", "", "", "", ""]),
];
debt.getRange("H8:H16").formulas = Array.from({ length: 9 }, (_, i) => [`=IF(B${i + 8}=\"\",\"\",E${i + 8}+F${i + 8})`]);
header(debt, "B7:I7");
debt.getRange("C8:C16").format.numberFormat = currency;
debt.getRange("D8:D16").format.numberFormat = "0.00%";
debt.getRange("E8:F16").format.numberFormat = currency;
debt.getRange("H8:H16").format.numberFormat = currency;
card(debt, "B7:I16", theme.white);
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
section(debt, "J15:K15", "Daily Log reminder", theme.mint);
debt.getRange("J16:K18").values = [
  ["When a debt payment actually happens, enter it on Daily Log with Category = Debt.", ""],
  ["When a subscription charge actually happens, enter it on Daily Log with Category = Subscriptions.", ""],
  ["This keeps budget and actuals clean without double-entry.", ""],
];
merge(debt, ["J16:K16", "J17:K17", "J18:K18"]);
card(debt, "J16:K18", theme.mint);

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
  ["What category is most over budget?", ""],
  ["What transaction caused the leak?", ""],
  ["Was it a need, timing issue, or convenience purchase?", ""],
  ["What subscription will I cut, pause, or justify?", ""],
  ["How much will I move to savings next payday?", ""],
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
action.getRange("B18").formulas = [["='Paycheck Map'!K14"]];
merge(action, ["B18:K18", "B19:K19"]);
action.getRange("B18:K18").format.font = { bold: true, color: theme.green, size: 16 };
card(action, "B18:K19", theme.mint);
section(action, "B22:K22", "The 20-minute monthly rhythm");
action.getRange("B23:K28").values = [
  ["1", "Set or adjust the budget on Paycheck Map.", "", "", "", "", "", "", "", ""],
  ["2", "Enter actual income and spending in Daily Log.", "", "", "", "", "", "", "", ""],
  ["3", "Check bills due before each paycheck.", "", "", "", "", "", "", "", ""],
  ["4", "Review subscriptions and debt pressure.", "", "", "", "", "", "", "", ""],
  ["5", "Pick one money leak to fix.", "", "", "", "", "", "", "", ""],
  ["6", "Return to the dashboard and write the next move.", "", "", "", "", "", "", "", ""],
];
merge(action, ["C23:K23", "C24:K24", "C25:K25", "C26:K26", "C27:K27", "C28:K28"]);
action.getRange("B23:B28").format.font.bold = true;
card(action, "B23:K28", theme.cream);

for (const sheet of [start, dash, map, log, bills, debt, action]) {
  sheet.getRange("B2:K4").format.borders = { preset: "outside", style: "thin", color: theme.green };
  sheet.getRange("B6:K80").format.borders = { preset: "inside", style: "thin", color: "#EFE7DA" };
  sheet.getRange("B6:K80").format.borders = { preset: "outside", style: "thin", color: theme.line };
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
  ["Start Here", "B2:K24"],
  ["Dashboard", "B2:K31"],
  ["Paycheck Map", "B2:K27"],
  ["Daily Log", "B2:L51"],
  ["Bills & Timing", "B2:K33"],
  ["Debt & Subscriptions", "B2:K33"],
  ["Action Plan", "B2:K28"],
]) {
  await workbook.render({ sheetName, range, scale: 1 });
}

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(new URL("money-leak-finder.xlsx", outputDir));
