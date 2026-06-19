import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL("../outputs/profit-hunter-money-leak-finder/", import.meta.url);
const outputDirPath = fileURLToPath(outputDir);
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
const excelDate = (year, month, day) => Math.floor((Date.UTC(year, month, day) - Date.UTC(1899, 11, 30)) / 86400000);

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
  M: 150,
  N: 150,
  O: 150,
  P: 150,
  Q: 150,
  R: 150,
  S: 150,
  T: 150,
  U: 150,
  V: 150,
  W: 150,
};

const addSheet = (name) => {
  const sheet = workbook.worksheets.add(name);
  Object.entries(widths).forEach(([col, px]) => {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
  });
  sheet.getRange("A1:W90").format.font.name = "Aptos";
  sheet.getRange("A1:W90").format.font.size = 11;
  sheet.getRange("A1:W90").format.wrapText = true;
  sheet.getRange("A1:W90").format.font.color = theme.ink;
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

const expenseCategories = [
  ["Housing", 1850, "Rent, mortgage, HOA dues, or property fees."],
  ["Utilities", 250, "Electricity, water, gas, trash, and similar household utilities."],
  ["Insurance", 240, "Auto, renters, homeowners, life, or other insurance premiums."],
  ["Phone & Internet", 170, "Mobile phone, home internet, and related communication costs."],
  ["Loan / Debt Payment", 520, "Car loans, student loans, personal loans, or other non-credit-card debt payments."],
  ["Credit Card Payments", 300, "Payments from checking to a credit card. Do not count as spending if card purchases are logged by category."],
  ["Subscriptions", 218, "Recurring apps, software, memberships, and streaming services."],
  ["Savings Transfers", 400, "Money moved to emergency savings, sinking funds, or investments."],
  ["Transfers", 0, "Money moved between your own checking, savings, or other accounts."],
  ["Groceries", 500, "Food at home and normal household basics."],
  ["Transportation", 150, "Gas, parking, rideshare, tolls, transit, and routine vehicle costs."],
  ["Medical / Health", 100, "Copays, prescriptions, therapy, dental, vision, and health supplies."],
  ["Childcare / Family", 0, "Daycare, babysitting, school costs, child support, or family care."],
  ["Pets", 90, "Pet food, supplies, grooming, boarding, and veterinary costs."],
  ["Food & Dining", 250, "Restaurants, coffee, airport sandwiches, and convenience meals."],
  ["Shopping / Personal", 200, "Clothes, personal care, household shopping, and everyday wants."],
  ["Entertainment", 100, "Movies, hobbies, games, activities, and local events."],
  ["Travel", 125, "Flights, lodging, rental cars, and trip-related spending."],
  ["Irregular Planned", 250, "Expected but non-monthly costs such as gifts, annual fees, repairs, and holidays."],
  ["Miscellaneous", 50, "Truly uncategorized small spending. Review this bucket if it grows."],
];

const typeForCategory = (category) =>
  category === "Credit Card Payments"
    ? "Credit Card Payment"
    : category === "Savings Transfers"
      ? "Savings Transfer"
      : category === "Transfers"
        ? "Transfer"
        : "Expense";

const planRows = [
  ["Income", "Income", 4850, "All paychecks and cash received this month."],
  ...expenseCategories.map(([category, budget, description]) => [category, typeForCategory(category), budget, description]),
  ["Current Emergency Savings", "Info", "", "Enter the cash you currently have set aside for emergencies."],
  ["Starter Emergency Target", "Info", "", "A simple first milestone before getting aggressive elsewhere."],
];

const logRows = [
  [excelDate(2026, 0, 2), "January paychecks", "Income", 4800, "Income", "", "Example income entry"],
  [excelDate(2026, 0, 3), "Rent / mortgage", "Housing", 1850, "Expense", "Need", ""],
  [excelDate(2026, 0, 5), "Utilities and insurance", "Utilities", 350, "Expense", "Need", ""],
  [excelDate(2026, 0, 8), "Groceries", "Groceries", 465, "Expense", "Need", ""],
  [excelDate(2026, 0, 12), "Transportation", "Transportation", 138, "Expense", "Need", ""],
  [excelDate(2026, 0, 18), "Dining and coffee", "Food & Dining", 285, "Expense", "Want", ""],
  [excelDate(2026, 1, 2), "February paychecks", "Income", 4950, "Income", "", "Example income entry"],
  [excelDate(2026, 1, 3), "Rent / mortgage", "Housing", 1850, "Expense", "Need", ""],
  [excelDate(2026, 1, 5), "Utilities", "Utilities", 265, "Expense", "Need", ""],
  [excelDate(2026, 1, 7), "Insurance", "Insurance", 240, "Expense", "Need", ""],
  [excelDate(2026, 1, 9), "Groceries", "Groceries", 510, "Expense", "Need", ""],
  [excelDate(2026, 1, 13), "Dog food and supplies", "Pets", 92, "Expense", "Need", ""],
  [excelDate(2026, 1, 20), "Car loan and student loan", "Loan / Debt Payment", 520, "Expense", "Need", ""],
  [excelDate(2026, 1, 22), "Credit card payment", "Credit Card Payments", 200, "Credit Card Payment", "", "Bank payment only; card purchases should be logged by spending category"],
  [excelDate(2026, 1, 25), "Checking to emergency savings", "Savings Transfers", 250, "Savings Transfer", "", "Transfer to savings; not ordinary spending"],
  [excelDate(2026, 2, 2), "March paychecks", "Income", 4875, "Income", "", "Example income entry"],
  [excelDate(2026, 2, 3), "Rent / mortgage", "Housing", 1850, "Expense", "Need", ""],
  [excelDate(2026, 2, 5), "Phone and internet", "Phone & Internet", 170, "Expense", "Need", ""],
  [excelDate(2026, 2, 11), "Gas and parking", "Transportation", 176, "Expense", "Need", ""],
  [excelDate(2026, 2, 15), "Airport sandwich", "Food & Dining", 18, "Expense", "Want", "Small leaks become visible here"],
  [excelDate(2026, 2, 21), "Streaming and apps", "Subscriptions", 218, "Expense", "Want", ""],
  [excelDate(2026, 5, 1), "Paycheck 1", "Income", 2450, "Income", "", "Example income entry"],
  [excelDate(2026, 5, 1), "Rent / mortgage", "Housing", 1850, "Expense", "Need", "Example fixed cost"],
  [excelDate(2026, 5, 3), "Grocery run", "Groceries", 112, "Expense", "Need", ""],
  [excelDate(2026, 5, 5), "Dog food", "Pets", 48, "Expense", "Need", "This is how normal real-life spending gets captured"],
  [excelDate(2026, 5, 8), "Airport sandwich", "Food & Dining", 18, "Expense", "Want", "Small leaks become visible here"],
  [excelDate(2026, 5, 10), "Streaming bundle", "Subscriptions", 39, "Expense", "Want", ""],
  [excelDate(2026, 5, 12), "Credit card payment", "Credit Card Payments", 200, "Credit Card Payment", "", "Do not also categorize this as groceries/dining/etc."],
  [excelDate(2026, 5, 13), "Savings transfer", "Savings Transfers", 150, "Savings Transfer", "", "Movement to savings, not spending"],
  [excelDate(2026, 5, 15), "Paycheck 2", "Income", 2400, "Income", "", "Example income entry"],
];

const monthRows = Array.from({ length: 12 }, (_, index) => [
  excelDate(2026, index, 1),
  4850,
  ...expenseCategories.map(([, budget]) => budget),
]);

const billsRows = [
  [1, "Rent / mortgage", 1850, "Housing", "Yes", "Paycheck 1", "Largest fixed cost"],
  [5, "Internet", 75, "Phone & Internet", "Yes", "Paycheck 1", ""],
  [8, "Car payment", 420, "Loan / Debt Payment", "Yes", "Paycheck 1", ""],
  [12, "Phone", 95, "Phone & Internet", "Yes", "Paycheck 1", ""],
  [15, "Utilities", 190, "Utilities", "No", "Paycheck 2", "Average monthly amount"],
  [18, "Insurance", 160, "Insurance", "Yes", "Paycheck 2", ""],
  [22, "Credit card payment", 210, "Credit Card Payments", "Yes", "Paycheck 2", "Bank payment to card, not purchase detail"],
  [25, "Student loan", 310, "Loan / Debt Payment", "Yes", "Paycheck 2", ""],
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
  ...expenseCategories.map(([category, , description]) => [category, description, typeForCategory(category)]),
];

const start = addSheet("Start Here");
titleBlock(
  start,
  "Money Leak Finder",
  "A guided paycheck audit for people who earn enough but still feel squeezed.",
  "Start here: follow the tabs from left to right."
);
section(start, "B6:K6", "How this workbook works");
start.getRange("B7:K14").values = [
  ["1. Start Here", "Read the flow and what each tab is for.", "", "", "", "", "", "", "", ""],
  ["2. Dashboard", "Choose any start and end date to review one month, several months, or year-to-date.", "", "", "", "", "", "", "", ""],
  ["3. Monthly Plan", "Store a separate budget for every month so prior months never disappear.", "", "", "", "", "", "", "", ""],
  ["4. Paycheck Map", "Select one month for a focused budget-vs-actual review.", "", "", "", "", "", "", "", ""],
  ["5. Daily Log", "Keep adding real income, spending, card payments, and transfers. Dates drive every report automatically.", "", "", "", "", "", "", "", ""],
  ["6. Trends", "Compare monthly income, spending, flex cash, and category percentages across the year.", "", "", "", "", "", "", "", ""],
  ["7. Bills & Timing", "Plan due dates so you can spot paycheck timing traps.", "", "", "", "", "", "", "", ""],
  ["8. Debt & Subscriptions", "Enter debt details clearly and choose what to cut, pause, or attack.", "", "", "", "", "", "", "", ""],
];
merge(start, ["C7:K7", "C8:K8", "C9:K9", "C10:K10", "C11:K11", "C12:K12", "C13:K13", "C14:K14"]);
start.getRange("B7:B14").format.font.bold = true;
card(start, "B7:K14", theme.white);
section(start, "B16:K16", "What to do next");
start.getRange("B17:K21").values = [
  ["Step 1", "Enter a budget for each month on Monthly Plan.", "", "", "", "", "", "", "", ""],
  ["Step 2", "Keep adding transactions to Daily Log. Purchases, credit card payments, and transfers each have their own categories.", "", "", "", "", "", "", "", ""],
  ["Step 3", "Use Paycheck Map for one-month budget-vs-actual review.", "", "", "", "", "", "", "", ""],
  ["Step 4", "Use Dashboard and Trends for multi-month and yearly analysis.", "", "", "", "", "", "", "", ""],
  ["Step 5", "Use Bills, Debt, and Action Plan to decide what changes next.", "", "", "", "", "", "", "", ""],
];
merge(start, ["C17:K17", "C18:K18", "C19:K19", "C20:K20", "C21:K21"]);
start.getRange("B17:B21").format.font.bold = true;
card(start, "B17:K21", theme.mint);
section(start, "B23:K23", "Important note");
start.getRange("B24:K25").values = [
  ["This workbook is educational. It is meant to show cash-flow patterns, not provide individualized financial, legal, or investment advice.", "", "", "", "", "", "", "", "", ""],
  ["The daily log is where the trust is built: if spending gets entered there, the dashboard tells the truth.", "", "", "", "", "", "", "", "", ""],
];
merge(start, ["B24:K24", "B25:K25"]);
card(start, "B24:K25", theme.cream);

const dash = addSheet("Dashboard");
titleBlock(
  dash,
  "Dashboard",
  "Choose any reporting period and see where the money went.",
  "Change the start and end dates to review one month, several months, or year-to-date."
);
section(dash, "B6:K6", "Reporting Period");
dash.getRange("B7:G8").values = [
  ["Start Date", excelDate(2026, 0, 1), "End Date", excelDate(2026, 5, 30), "Quick use", "Change these two dates"],
  ["Examples", "Jan 1 to Jan 31", "", "Jan 1 to Jun 30", "", "Any custom period works"],
];
dash.getRange("C7:E7").format.fill = theme.cream;
dash.getRange("C7:E7").format.font = { color: "#0000FF", bold: true };
dash.getRange("C7:E7").format.numberFormat = "mmm d, yyyy";
card(dash, "B7:G8", theme.white);

section(dash, "B10:K10", "Money Snapshot");
dash.getRange("B11:C13").values = [["Actual income", ""], ["", ""], ["", ""]];
dash.getRange("D11:E13").values = [["Actual spending", ""], ["", ""], ["", ""]];
dash.getRange("F11:G13").values = [["True flex cash", ""], ["", ""], ["", ""]];
dash.getRange("H11:I13").values = [["Cash movements", ""], ["", ""], ["", ""]];
dash.getRange("J11:K13").values = [["Largest expense", ""], ["", ""], ["", ""]];
merge(dash, ["B11:C11", "B12:C13", "D11:E11", "D12:E13", "F11:G11", "F12:G13", "H11:I11", "H12:I13", "J11:K11", "J12:K13"]);
dash.getRange("B12").formulas = [["=SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Income\",'Daily Log'!B:B,\">=\"&C7,'Daily Log'!B:B,\"<=\"&E7)"]];
dash.getRange("D12").formulas = [["=SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Expense\",'Daily Log'!B:B,\">=\"&C7,'Daily Log'!B:B,\"<=\"&E7)"]];
dash.getRange("F12").formulas = [["=B12-D12-H12"]];
dash.getRange("H12").formulas = [["=SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Savings Transfer\",'Daily Log'!B:B,\">=\"&C7,'Daily Log'!B:B,\"<=\"&E7)+SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Credit Card Payment\",'Daily Log'!B:B,\">=\"&C7,'Daily Log'!B:B,\"<=\"&E7)+SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Transfer\",'Daily Log'!B:B,\">=\"&C7,'Daily Log'!B:B,\"<=\"&E7)"]];
dash.getRange("J12").formulas = [["=INDEX(B18:B37,MATCH(MAXIFS(D18:D37,B18:B37,\"<>Savings Transfers\",B18:B37,\"<>Credit Card Payments\",B18:B37,\"<>Transfers\"),D18:D37,0))"]];
card(dash, "B11:C13", theme.mint);
card(dash, "D11:E13", theme.sky);
card(dash, "F11:G13", theme.cream);
card(dash, "H11:I13", theme.warm);
card(dash, "J11:K13", theme.rose);
dash.getRange("B11:K11").format.font = { bold: true, color: theme.muted, size: 11 };
dash.getRange("B12:G13").format.font = { bold: true, color: theme.green, size: 22 };
dash.getRange("B12:G13").format.numberFormat = currency;
dash.getRange("H12:I13").format.font = { bold: true, color: theme.green, size: 22 };
dash.getRange("H12:I13").format.numberFormat = currency;
dash.getRange("J12:K13").format.font = { bold: true, color: theme.rust, size: 14 };

section(dash, "B16:G16", "Category Analysis");
dash.getRange("B17:G37").values = [
  ["Category", "Budget", "Actual", "% of Income", "Variance", "Signal"],
  ...expenseCategories.map(([category]) => [category, "", "", "", "", ""]),
];
const monthlyPlanCols = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"];
dash.getRange("C18:C28").formulas = monthlyPlanCols.map((col) => [`=SUMIFS('Monthly Plan'!${col}$8:${col}$19,'Monthly Plan'!$B$8:$B$19,\">=\"&DATE(YEAR($C$7),MONTH($C$7),1),'Monthly Plan'!$B$8:$B$19,\"<=\"&EOMONTH($E$7,0))`]);
dash.getRange("D18:D28").formulas = Array.from({ length: 11 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!D:D,B${i + 18},'Daily Log'!B:B,\">=\"&$C$7,'Daily Log'!B:B,\"<=\"&$E$7)`]);
dash.getRange("E18:E28").formulas = Array.from({ length: 11 }, (_, i) => [`=IF($B$12=0,0,D${i + 18}/$B$12)`]);
dash.getRange("F18:F28").formulas = Array.from({ length: 11 }, (_, i) => [`=C${i + 18}-D${i + 18}`]);
dash.getRange("G18:G28").formulas = Array.from({ length: 11 }, (_, i) => [`=IF(F${i + 18}<0,\"Over budget\",\"OK\")`]);
header(dash, "B17:G17");
dash.getRange("C18:D28").format.numberFormat = currency;
dash.getRange("E18:E28").format.numberFormat = percent;
dash.getRange("F18:F28").format.numberFormat = currency;
card(dash, "B17:G28", theme.white);
dash.getRange("E18:E28").conditionalFormats.add("dataBar", { color: theme.blue, gradient: true });
const expandedDashboardCols = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W"];
dash.getRange("C18:C37").formulas = expandedDashboardCols.map((col) => [`=SUMIFS('Monthly Plan'!${col}$8:${col}$19,'Monthly Plan'!$B$8:$B$19,">="&DATE(YEAR($C$7),MONTH($C$7),1),'Monthly Plan'!$B$8:$B$19,"<="&EOMONTH($E$7,0))`]);
dash.getRange("D18:D37").formulas = Array.from({ length: 20 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!D:D,B${i + 18},'Daily Log'!B:B,">="&$C$7,'Daily Log'!B:B,"<="&$E$7)`]);
dash.getRange("E18:E37").formulas = Array.from({ length: 20 }, (_, i) => [`=IF($B$12=0,0,D${i + 18}/$B$12)`]);
dash.getRange("F18:F37").formulas = Array.from({ length: 20 }, (_, i) => [`=C${i + 18}-D${i + 18}`]);
dash.getRange("G18:G37").formulas = Array.from({ length: 20 }, (_, i) => [`=IF(OR(B${i + 18}="Credit Card Payments",B${i + 18}="Savings Transfers",B${i + 18}="Transfers"),"Cash movement",IF(F${i + 18}<0,"Over budget","OK"))`]);
dash.getRange("C18:D37").format.numberFormat = currency;
dash.getRange("E18:E37").format.numberFormat = percent;
dash.getRange("F18:F37").format.numberFormat = currency;
card(dash, "B17:G37", theme.white);
dash.getRange("E18:E37").conditionalFormats.add("dataBar", { color: theme.blue, gradient: true });

section(dash, "I16:K16", "Period Insight");
dash.getRange("I17:K23").values = [
  ["Biggest expense category", "", ""],
  ["", "", ""],
  ["Share of income", "", ""],
  ["", "", ""],
  ["Budget variance", "", ""],
  ["", "", ""],
  ["Tip", "", ""],
];
dash.getRange("I18").formulas = [["=J12"]];
dash.getRange("I20").formulas = [["=INDEX(E18:E37,MATCH(J12,B18:B37,0))"]];
dash.getRange("I22").formulas = [["=SUM(F18:F37)"]];
dash.getRange("I24:K25").values = [["Use the date boxes above to compare a month, quarter, or year-to-date without changing the Daily Log.", "", ""], ["", "", ""]];
merge(dash, ["I17:K17", "I18:K18", "I19:K19", "I20:K20", "I21:K21", "I22:K22", "I23:K23", "I24:K25"]);
dash.getRange("I17:I23").format.font.bold = true;
dash.getRange("I18:K18").format.font = { bold: true, color: theme.rust, size: 15 };
dash.getRange("I20:K20").format.numberFormat = percent;
dash.getRange("I22:K22").format.numberFormat = currency;
card(dash, "I17:K25", theme.cream);

section(dash, "B38:K38", "Mini money meeting");
dash.getRange("B39:K42").values = [
  ["1", "What category is over budget and why?", "", "", "", "", "", "", "", ""],
  ["2", "Was the issue a real need, a timing problem, or a convenience purchase?", "", "", "", "", "", "", "", ""],
  ["3", "What purchase pattern would I change before next payday?", "", "", "", "", "", "", "", ""],
  ["4", "What one move will make next payday calmer?", "", "", "", "", "", "", "", ""],
];
merge(dash, ["C39:K39", "C40:K40", "C41:K41", "C42:K42"]);
dash.getRange("B39:B42").format.font.bold = true;
card(dash, "B39:K42", theme.mint);

const monthly = addSheet("Monthly Plan");
titleBlock(
  monthly,
  "Monthly Plan",
  "Keep a separate budget for every month of the year.",
  "Change the blue input cells. Prior months remain available for comparisons."
);
section(monthly, "B6:W6", "2026 Monthly Budgets - type only in the blue cells");
monthly.getRange("B7:W19").values = [
  ["Month", "Expected Income", ...expenseCategories.map(([category]) => category)],
  ...monthRows,
];
header(monthly, "B7:W7");
monthly.getRange("B8:B19").format.numberFormat = "mmm yyyy";
monthly.getRange("C8:W19").format.numberFormat = currency;
monthly.getRange("B8:B19").format.fill = theme.paper;
monthly.getRange("C8:W19").format.fill = theme.cream;
monthly.getRange("C8:W19").format.font = { color: "#0000FF" };
card(monthly, "B7:W19", theme.white);
section(monthly, "B22:W22", "What every column means");
monthly.getRange("B23:W27").values = [
  ["Column B = the month. Column C = the income you expect to receive during that month.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Columns D through W = the amount you plan to spend or move in the category named at the top of that column.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Do not enter actual spending on this tab. Enter real transactions on Daily Log; Paycheck Map and Dashboard calculate the actuals.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Credit Card Payments and Transfers are cash movement categories. Use them to track money moving between accounts or to a card without double-counting spending.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Irregular Planned means expected but non-monthly costs. Miscellaneous is only for genuinely uncategorized small spending.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
];
merge(monthly, ["B23:W23", "B24:W24", "B25:W25", "B26:W26", "B27:W27"]);
card(monthly, "B23:W27", theme.mint);

const map = addSheet("Paycheck Map");
titleBlock(
  map,
  "Paycheck Map",
  "Choose one month for a focused budget-vs-actual review.",
  "Budget comes from Monthly Plan. Actuals come from Daily Log."
);
section(map, "B6:H6", "Selected Month");
map.getRange("B7:D8").values = [["Month to review", excelDate(2026, 5, 1), "Change this date"], ["Report covers", "", ""]];
map.getRange("C7").format.fill = theme.cream;
map.getRange("C7").format.font = { color: "#0000FF", bold: true };
map.getRange("C7").format.numberFormat = "mmm yyyy";
map.getRange("C8").formulas = [["=TEXT(C7,\"mmm d\")&\" to \"&TEXT(EOMONTH(C7,0),\"mmm d, yyyy\")"]];
card(map, "B7:D8", theme.white);
section(map, "B10:H10", "Budget vs Actual");
map.getRange("B11:H34").values = [
  ["Category", "Type", "Budget (auto)", "Actual (auto)", "Difference (auto)", "Signal", "What this means"],
  ...planRows.map(([category, type, , description]) => [category, type, "", "", "", "", description]),
];
header(map, "B11:H11");
const budgetCols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"];
map.getRange("D12:D23").formulas = budgetCols.map((col) => [`=SUMIF('Monthly Plan'!B:B,$C$7,'Monthly Plan'!${col}:${col})`]);
map.getRange("E12").formulas = [["=SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Income\",'Daily Log'!B:B,\">=\"&$C$7,'Daily Log'!B:B,\"<=\"&EOMONTH($C$7,0))"]];
map.getRange("E13:E23").formulas = Array.from({ length: 11 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!D:D,B${i + 13},'Daily Log'!B:B,\">=\"&$C$7,'Daily Log'!B:B,\"<=\"&EOMONTH($C$7,0))`]);
map.getRange("E24:E25").values = [[700], [1000]];
map.getRange("F12:F23").formulas = Array.from({ length: 12 }, (_, i) => {
  const row = i + 12;
  return [`=IF(C${row}=\"Income\",E${row}-D${row},D${row}-E${row})`];
});
map.getRange("G12:G23").formulas = Array.from({ length: 12 }, (_, i) => {
  const row = i + 12;
  return [`=IF(C${row}=\"Income\",IF(F${row}<0,\"Under plan\",\"On track\"),IF(F${row}<0,\"Over budget\",\"OK\"))`];
});
map.getRange("F24:F25").formulas = [[""], ["=MAX(0,E25-E24)"]];
map.getRange("G24").values = [["Manual input"]];
map.getRange("G25").formulas = [["=IF(F25>0,\"Gap remains\",\"Funded\")"]];
map.getRange("D12:F25").format.numberFormat = currency;
card(map, "B11:H25", theme.white);
const expandedBudgetCols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W"];
map.getRange("D12:D32").formulas = expandedBudgetCols.map((col) => [`=SUMIF('Monthly Plan'!B:B,$C$7,'Monthly Plan'!${col}:${col})`]);
map.getRange("E13:E32").formulas = Array.from({ length: 20 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!D:D,B${i + 13},'Daily Log'!B:B,">="&$C$7,'Daily Log'!B:B,"<="&EOMONTH($C$7,0))`]);
map.getRange("E33:E34").values = [[700], [1000]];
map.getRange("F12:F32").formulas = Array.from({ length: 21 }, (_, i) => {
  const row = i + 12;
  return [`=IF(C${row}="Income",E${row}-D${row},D${row}-E${row})`];
});
map.getRange("G12:G32").formulas = Array.from({ length: 21 }, (_, i) => {
  const row = i + 12;
  return [`=IF(C${row}="Income",IF(F${row}<0,"Under plan","On track"),IF(OR(C${row}="Credit Card Payment",C${row}="Savings Transfer",C${row}="Transfer"),"Cash movement",IF(F${row}<0,"Over budget","OK")))`];
});
map.getRange("F33:F34").formulas = [[""], ["=MAX(0,E34-E33)"]];
map.getRange("G33").values = [["Manual input"]];
map.getRange("G34").formulas = [["=IF(F34>0,\"Gap remains\",\"Funded\")"]];
map.getRange("D12:F34").format.numberFormat = currency;
map.getRange("E33:E34").format = { fill: theme.cream, font: { color: "#0000FF", bold: true }, numberFormat: currency };
map.getRange("D12:G32").format.font.color = "#008000";
card(map, "B11:H34", theme.white);
section(map, "B35:H35", "Where to enter information");
map.getRange("B36:H40").values = [
  ["BLUE CELL C7: choose the month you want to review.", "", "", "", "", "", ""],
  ["Budget (auto): comes from Monthly Plan. Do not type a budget into this table.", "", "", "", "", "", ""],
  ["Actual (auto): comes from Daily Log. Enter purchases and paychecks there, not here.", "", "", "", "", "", ""],
  ["Credit card payments and transfers appear as cash movement rows. They do not count as ordinary spending on the Dashboard.", "", "", "", "", "", ""],
  ["BLUE CELLS E33:E34: enter your current emergency savings and target. These are balances, so they are manual.", "", "", "", "", "", ""],
];
merge(map, ["B36:H36", "B37:H37", "B38:H38", "B39:H39", "B40:H40"]);
card(map, "B36:H40", theme.mint);
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
  ["=E12"],
  ["=SUMIFS(E13:E32,C13:C32,\"Expense\")"],
  ["=E12-SUMIFS(E13:E32,C13:C32,\"Expense\")-SUMIFS(E13:E32,C13:C32,\"Credit Card Payment\")-SUMIFS(E13:E32,C13:C32,\"Savings Transfer\")-SUMIFS(E13:E32,C13:C32,\"Transfer\")"],
  ["=MAX(0,SUMIFS(E13:E32,C13:C32,\"Expense\")+SUMIFS(E13:E32,C13:C32,\"Credit Card Payment\")+SUMIFS(E13:E32,C13:C32,\"Savings Transfer\")+SUMIFS(E13:E32,C13:C32,\"Transfer\")-E12)"],
  ["=F34"],
];
map.getRange("K13").formulas = [["=IF(MIN(F13:F32)<0,INDEX(B13:B32,MATCH(MIN(F13:F32),F13:F32,0)),\"No over-budget category yet\")"]];
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
section(log, "B6:H6", "Transaction log - use the Category and Need / Want dropdowns");
log.getRange("B7:H79").values = [
  ["Date", "Description", "Category", "Amount", "Type", "Need / Want", "Notes"],
  ...logRows,
  ...Array.from({ length: 42 }, () => ["", "", "", "", "", "", ""]),
];
header(log, "B7:H7");
log.getRange("B8:B79").format.numberFormat = "yyyy-mm-dd";
log.getRange("E8:E79").format.numberFormat = currency;
log.getRange("B8:H79").format.fill = theme.cream;
log.getRange("B8:H79").format.font.color = "#0000FF";
card(log, "B7:H79", theme.white);
log.getRange("B8:H79").format.fill = theme.cream;
log.getRange("B8:H79").format.font.color = "#0000FF";
section(log, "J6:K6", "Category guide");
log.getRange("J7:L28").values = [["Category", "Use for", "Type"], ...categoryGuide];
header(log, "J7:L7");
card(log, "J7:L28", theme.cream);
log.getRange("D8:D79").dataValidation = {
  rule: { type: "list", formula1: "$J$8:$J$28" },
};
log.getRange("G8:G79").dataValidation = {
  rule: { type: "list", values: ["Need", "Want"] },
};
log.getRange("F8:F79").formulas = Array.from({ length: 72 }, (_, index) => {
  const row = index + 8;
  return [`=IF(D${row}="","",INDEX($L$8:$L$28,MATCH(D${row},$J$8:$J$28,0)))`];
});
log.getRange("F8:F79").format.fill = theme.mint;
log.getRange("F8:F79").format.font = { color: "#000000" };
section(log, "J29:L29", "What counts here");
log.getRange("J30:L32").values = [
  ["Log credit card purchases by what they actually were: groceries, dining, pets, travel, etc.", "", ""],
  ["Log the bank payment to the card as Credit Card Payments so the payment does not double-count as spending.", "", ""],
  ["Log checking-to-savings or savings-to-checking movement as Transfers or Savings Transfers, not as income or spending.", "", ""],
];
merge(log, ["J30:L30", "J31:L31", "J32:L32"]);
card(log, "J30:L32", theme.mint);
section(log, "J35:L35", "Daily Log controls", theme.sky);
log.getRange("J36:L39").values = [
  ["Category", "Choose from the dropdown; the choices match the Category Guide.", ""],
  ["Type", "Fills automatically as Income, Expense, Credit Card Payment, Savings Transfer, or Transfer.", ""],
  ["Need / Want", "Choose Need or Want from the dropdown. Leave blank for income.", ""],
  ["Color key", "Blue/yellow cells are inputs. Green cells calculate automatically.", ""],
];
merge(log, ["K36:L36", "K37:L37", "K38:L38", "K39:L39"]);
log.getRange("J36:J39").format.font.bold = true;
card(log, "J36:L39", theme.white);

const trends = addSheet("Trends");
titleBlock(
  trends,
  "Monthly Trends",
  "See the story across the year instead of judging one isolated month.",
  "Income, spending, flex cash, and category percentages update from Daily Log."
);
section(trends, "B6:G6", "12-Month Summary");
trends.getRange("B7:G19").values = [
  ["Month", "Income", "Spending", "True Flex Cash", "Savings", "Savings Rate"],
  ...Array.from({ length: 12 }, (_, index) => [excelDate(2026, index, 1), "", "", "", "", ""]),
];
trends.getRange("C8:C19").formulas = Array.from({ length: 12 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Income\",'Daily Log'!B:B,\">=\"&B${i + 8},'Daily Log'!B:B,\"<=\"&EOMONTH(B${i + 8},0))`]);
trends.getRange("D8:D19").formulas = Array.from({ length: 12 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!F:F,\"Expense\",'Daily Log'!B:B,\">=\"&B${i + 8},'Daily Log'!B:B,\"<=\"&EOMONTH(B${i + 8},0))`]);
trends.getRange("E8:E19").formulas = Array.from({ length: 12 }, (_, i) => [`=C${i + 8}-D${i + 8}`]);
trends.getRange("F8:F19").formulas = Array.from({ length: 12 }, (_, i) => [`=SUMIFS('Daily Log'!E:E,'Daily Log'!D:D,\"Savings Transfers\",'Daily Log'!B:B,\">=\"&B${i + 8},'Daily Log'!B:B,\"<=\"&EOMONTH(B${i + 8},0))`]);
trends.getRange("G8:G19").formulas = Array.from({ length: 12 }, (_, i) => [`=IF(C${i + 8}=0,0,F${i + 8}/C${i + 8})`]);
header(trends, "B7:G7");
trends.getRange("B8:B19").format.numberFormat = "mmm yyyy";
trends.getRange("C8:F19").format.numberFormat = currency;
trends.getRange("G8:G19").format.numberFormat = percent;
card(trends, "B7:G19", theme.white);

section(trends, "B22:G22", "Expense Categories as % of Income");
trends.getRange("B23:G35").values = [
  ["Month", "Housing %", "Transportation %", "Groceries %", "Food & Dining %", "Loan / Debt %"],
  ...Array.from({ length: 12 }, (_, index) => [excelDate(2026, index, 1), "", "", "", "", ""]),
];
for (const [col, category] of [["C", "Housing"], ["D", "Transportation"], ["E", "Groceries"], ["F", "Food & Dining"], ["G", "Loan / Debt Payment"]]) {
  trends.getRange(`${col}24:${col}35`).formulas = Array.from({ length: 12 }, (_, i) => [`=IF($C${i + 8}=0,0,SUMIFS('Daily Log'!E:E,'Daily Log'!D:D,\"${category}\",'Daily Log'!B:B,\">=\"&B${i + 24},'Daily Log'!B:B,\"<=\"&EOMONTH(B${i + 24},0))/$C${i + 8})`]);
}
header(trends, "B23:G23");
trends.getRange("B24:B35").format.numberFormat = "mmm yyyy";
trends.getRange("C24:G35").format.numberFormat = percent;
card(trends, "B23:G35", theme.white);
trends.getRange("C24:G35").conditionalFormats.add("dataBar", { color: theme.blue, gradient: true });

const trendChart = trends.charts.add("line", trends.getRange("B7:E19"));
trendChart.setPosition("I7", "P20");
trendChart.title = "Income, Spending, and Flex Cash by Month";
trendChart.hasLegend = true;
trendChart.xAxis = { axisType: "textAxis", tickLabelInterval: 1 };
trendChart.yAxis = { numberFormatCode: "$#,##0" };

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
bills.getRange("B29:K34").values = [
  ["This tab plans timing. Actual payments still belong in Daily Log using the Log Category shown above.", "", "", "", "", "", "", "", "", ""],
  ["Credit card payments are timing/cash movement. Credit card purchases should still be logged by their real category when you have the card detail.", "", "", "", "", "", "", "", "", ""],
  ["If most large bills hit before Paycheck 1 clears, timing may be the real problem.", "", "", "", "", "", "", "", "", ""],
  ["If possible, move one or two due dates so each paycheck carries a fair share.", "", "", "", "", "", "", "", "", ""],
  ["If a bill is on autopay, make sure the paycheck that covers it lands before the withdrawal.", "", "", "", "", "", "", "", "", ""],
  ["Next tab: Debt & Subscriptions. That is where the quiet monthly leaks usually show up.", "", "", "", "", "", "", "", "", ""],
];
merge(bills, ["B29:K29", "B30:K30", "B31:K31", "B32:K32", "B33:K33", "B34:K34"]);
card(bills, "B29:K34", theme.cream);

const debt = addSheet("Debt & Subscriptions");
titleBlock(
  debt,
  "Debt & Subscriptions",
  "Enter debt details clearly and choose one focus area.",
  "Purpose: reduce pressure without starving the rest of the month."
);
section(debt, "B6:I6", "Debt details - enter blue cells; Monthly Total calculates automatically");
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
debt.getRange("B8:G16").format.fill = theme.cream;
debt.getRange("B8:G16").format.font.color = "#0000FF";
debt.getRange("I8:I16").format.fill = theme.cream;
debt.getRange("I8:I16").format.font.color = "#0000FF";
debt.getRange("H8:H16").format.fill = theme.mint;
debt.getRange("H8:H16").format.font.color = "#000000";
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
debt.getRange("B21:D33").format.fill = theme.cream;
debt.getRange("B21:D33").format.font.color = "#0000FF";
debt.getRange("F21:G33").format.fill = theme.cream;
debt.getRange("F21:G33").format.font.color = "#0000FF";
debt.getRange("E21:E33").format.fill = theme.mint;
debt.getRange("E21:E33").format.font.color = "#000000";
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
  ["Car loans, student loans, and personal loans go to Daily Log with Category = Loan / Debt Payment.", ""],
  ["Credit card bank payments go to Category = Credit Card Payments so card purchase details are not double-counted.", ""],
  ["Subscriptions still go to Category = Subscriptions when the actual charge happens.", ""],
];
merge(debt, ["J16:K16", "J17:K17", "J18:K18"]);
card(debt, "J16:K18", theme.mint);
section(debt, "J20:K20", "Debt column guide", theme.sky);
debt.getRange("J21:K28").values = [
  ["Current Balance", "Total amount still owed today"],
  ["Interest Rate / APR", "Annual interest rate shown as a percentage"],
  ["Minimum Payment", "Required monthly payment"],
  ["Extra Target", "Optional extra amount you plan to pay"],
  ["Due Day", "Calendar day the payment is due"],
  ["Monthly Total", "Automatic: minimum payment plus extra target"],
  ["Notes", "Anything that helps explain the debt or strategy"],
  ["Color key", "Blue/yellow = enter it; green = calculated"],
];
debt.getRange("J21:J28").format.font.bold = true;
card(debt, "J21:K28", theme.white);

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
  ["1", "Set or adjust the budget on Monthly Plan.", "", "", "", "", "", "", "", ""],
  ["2", "Enter actual income and spending in Daily Log.", "", "", "", "", "", "", "", ""],
  ["3", "Check bills due before each paycheck.", "", "", "", "", "", "", "", ""],
  ["4", "Review subscriptions and debt pressure.", "", "", "", "", "", "", "", ""],
  ["5", "Pick one money leak to fix.", "", "", "", "", "", "", "", ""],
  ["6", "Return to the dashboard and write the next move.", "", "", "", "", "", "", "", ""],
];
merge(action, ["C23:K23", "C24:K24", "C25:K25", "C26:K26", "C27:K27", "C28:K28"]);
action.getRange("B23:B28").format.font.bold = true;
card(action, "B23:K28", theme.cream);

for (const sheet of [start, dash, monthly, map, log, trends, bills, debt, action]) {
  sheet.getRange("B2:K4").format.borders = { preset: "outside", style: "thin", color: theme.green };
  sheet.getRange("B6:W90").format.borders = { preset: "inside", style: "thin", color: "#EFE7DA" };
  sheet.getRange("B6:W90").format.borders = { preset: "outside", style: "thin", color: theme.line };
  sheet.getRange("A:A").format.columnWidthPx = 24;
}

for (const [sheet, range] of [
  [dash, "B17:G17"],
  [monthly, "B7:W7"],
  [map, "B11:H11"],
  [log, "B7:H7"],
  [log, "J7:L7"],
  [trends, "B7:G7"],
  [trends, "B23:G23"],
  [bills, "B7:H7"],
  [debt, "B7:I7"],
  [debt, "B20:G20"],
  [action, "B7:K7"],
]) {
  header(sheet, range);
}
monthly.getRange("B7:W7").format.rowHeightPx = 54;
map.getRange("B11:H11").format.rowHeightPx = 42;
debt.getRange("B7:I7").format.rowHeightPx = 42;
debt.getRange("B20:G20").format.rowHeightPx = 42;

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const [sheetName, range] of [
  ["Start Here", "B2:K25"],
  ["Dashboard", "B2:K42"],
  ["Monthly Plan", "B2:W27"],
  ["Paycheck Map", "B2:K40"],
  ["Daily Log", "B2:L40"],
  ["Trends", "B2:P35"],
  ["Bills & Timing", "B2:K34"],
  ["Debt & Subscriptions", "B2:K33"],
  ["Action Plan", "B2:K28"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1 });
  const previewName = sheetName.toLowerCase().replaceAll(" ", "-").replaceAll("&", "and");
  await fs.mkdir(outputDirPath, { recursive: true });
  await fs.writeFile(new URL(`preview-${previewName}.png`, outputDir), new Uint8Array(await preview.arrayBuffer()));
}

await fs.mkdir(outputDirPath, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(fileURLToPath(new URL("money-leak-finder.xlsx", outputDir)));
