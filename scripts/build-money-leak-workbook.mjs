import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL("../outputs/profit-hunter-money-leak-finder/", import.meta.url);
const workbook = Workbook.create();

const setupSheet = (sheet, widths = []) => {
  widths.forEach(([col, px]) => {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
  });
};

const dashboard = workbook.worksheets.add("Dashboard");
setupSheet(dashboard, [
  ["A", 30],
  ["B", 250],
  ["C", 145],
  ["D", 40],
  ["E", 235],
  ["F", 150],
]);

dashboard.getRange("B2:F2").values = [["Money Leak Finder", "", "", "", ""]];
dashboard.getRange("B3:F3").values = [["A quick paycheck audit for people who earn enough but still feel squeezed.", "", "", "", ""]];
dashboard.getRange("B5:C5").values = [["Input", "Monthly Amount"]];
dashboard.getRange("B6:C15").values = [
  ["Take-home pay", 4850],
  ["Fixed bills", 2940],
  ["Debt minimums", 520],
  ["Subscriptions", 218],
  ["Target savings transfer", 400],
  ["Groceries / gas / essentials", 650],
  ["Planned fun money", 300],
  ["Irregular expense buffer", 250],
  ["Current emergency savings", 700],
  ["Starter emergency target", 1000],
];
dashboard.getRange("C6:C15").format.numberFormat = "$#,##0";

dashboard.getRange("E5:F5").values = [["Output", "Amount"]];
dashboard.getRange("E6:E14").values = [
  ["Total committed money"],
  ["Money leak / overspend risk"],
  ["True flex cash"],
  ["Emergency fund gap"],
  ["Bills as % of take-home"],
  ["Debt as % of take-home"],
  ["Subscriptions as % of take-home"],
  ["Safety transfer as % of take-home"],
  ["Recommended first move"],
];
dashboard.getRange("F6:F13").formulas = [
  ["=SUM(C7:C13)"],
  ["=MAX(0,F6-C6)"],
  ["=MAX(0,C6-F6)"],
  ["=MAX(0,C15-C14)"],
  ["=C7/C6"],
  ["=C8/C6"],
  ["=C9/C6"],
  ["=C10/C6"],
];
dashboard.getRange("F14").formulas = [["=IF(F7>0,\"Cut or reschedule $\"&TEXT(F7,\"#,##0\"),IF(F9>0,\"Move $\"&TEXT(MIN(F8,F9),\"#,##0\")&\" toward starter emergency fund\",\"Protect this rhythm\"))"]];
dashboard.getRange("F6:F9").format.numberFormat = "$#,##0";
dashboard.getRange("F10:F13").format.numberFormat = "0%";

dashboard.getRange("B18:F18").values = [["How to read this", "", "", "", ""]];
dashboard.getRange("B19:F22").values = [
  ["Committed money is what already has a job before the month gets flexible.", "", "", "", ""],
  ["True flex cash is the number you can use for choices without pretending bills are optional.", "", "", "", ""],
  ["If the leak number is above zero, the paycheck is overbooked before the month starts.", "", "", "", ""],
  ["This is educational planning, not individualized financial advice.", "", "", "", ""],
];

for (const range of ["B2:F2", "B3:F3", "B18:F18", "B19:F19", "B20:F20", "B21:F21", "B22:F22"]) {
  dashboard.getRange(range).merge();
}

const bills = workbook.worksheets.add("Bill Calendar");
setupSheet(bills, [
  ["A", 110],
  ["B", 230],
  ["C", 130],
  ["D", 150],
  ["E", 160],
  ["F", 220],
]);
bills.getRange("A1:F1").values = [["Due Day", "Bill", "Amount", "Autopay?", "Paycheck Covered By", "Notes"]];
bills.getRange("A2:F15").values = [
  [1, "Rent / mortgage", 1850, "Yes", "Paycheck 1", "Largest fixed bill"],
  [5, "Internet", 75, "Yes", "Paycheck 1", ""],
  [8, "Car payment", 420, "Yes", "Paycheck 1", ""],
  [12, "Phone", 95, "Yes", "Paycheck 1", ""],
  [15, "Utilities", 190, "No", "Paycheck 2", "Average amount"],
  [18, "Insurance", 160, "Yes", "Paycheck 2", ""],
  [22, "Credit card minimum", 210, "Yes", "Paycheck 2", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
];
bills.getRange("C2:C15").format.numberFormat = "$#,##0";

const subscriptions = workbook.worksheets.add("Subscription Audit");
setupSheet(subscriptions, [
  ["A", 220],
  ["B", 130],
  ["C", 150],
  ["D", 170],
  ["E", 260],
]);
subscriptions.getRange("A1:E1").values = [["Subscription", "Monthly Cost", "Keep / Cut?", "Annual Cost", "Reason"]];
subscriptions.getRange("A2:E13").values = [
  ["Streaming bundle", 39, "Review", "", "Overlaps with another subscription"],
  ["Music", 11, "Keep", "", "Used daily"],
  ["Fitness app", 19, "Cut", "", "Not used this month"],
  ["Cloud storage", 10, "Keep", "", "Needed"],
  ["Meal plan app", 29, "Review", "", "Could pause"],
  ["News app", 7, "Cut", "", "Rarely used"],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
];
subscriptions.getRange("D2:D13").formulas = Array.from({ length: 12 }, (_, i) => [`=IF(B${i + 2}=\"\",\"\",B${i + 2}*12)`]);
subscriptions.getRange("B2:D13").format.numberFormat = "$#,##0";

const checklist = workbook.worksheets.add("Payday Checklist");
setupSheet(checklist, [
  ["A", 80],
  ["B", 560],
  ["C", 160],
]);
checklist.getRange("A1:C1").values = [["Done?", "Task", "Timing"]];
checklist.getRange("A2:C12").values = [
  ["", "Check which bills are due before the next paycheck", "Every payday"],
  ["", "Move fixed bill money first", "Every payday"],
  ["", "Pay debt minimums before extra payoff", "Every payday"],
  ["", "Move savings transfer before flexible spending", "Every payday"],
  ["", "Set the grocery / gas boundary", "Every payday"],
  ["", "Choose one subscription to keep, cut, or pause", "Monthly"],
  ["", "Add one irregular expense to the buffer", "Monthly"],
  ["", "Review true flex cash before weekend spending", "Weekly"],
  ["", "Update emergency fund progress", "Monthly"],
  ["", "Run a 20-minute money meeting", "Monthly"],
  ["", "Adjust bill due dates if timing is creating stress", "As needed"],
];

for (const sheet of [dashboard, bills, subscriptions, checklist]) {
  sheet.getRange("A1:Z200").format.font.name = "Aptos";
  sheet.getRange("A1:Z200").format.font.size = 11;
  sheet.getRange("A1:Z200").format.wrapText = true;
}

dashboard.getRange("B2:F2").format.font.size = 22;
dashboard.getRange("B2:F2").format.font.bold = true;
dashboard.getRange("B2:F2").format.font.color = "#193D34";
dashboard.getRange("B5:C5").format.font.bold = true;
dashboard.getRange("E5:F5").format.font.bold = true;
dashboard.getRange("B18:F18").format.font.bold = true;
bills.getRange("A1:F1").format.font.bold = true;
subscriptions.getRange("A1:E1").format.font.bold = true;
checklist.getRange("A1:C1").format.font.bold = true;

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "Dashboard", range: "B2:F22", scale: 1 });
await workbook.render({ sheetName: "Bill Calendar", range: "A1:F15", scale: 1 });
await workbook.render({ sheetName: "Subscription Audit", range: "A1:E13", scale: 1 });
await workbook.render({ sheetName: "Payday Checklist", range: "A1:C12", scale: 1 });

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(new URL("money-leak-finder.xlsx", outputDir));
