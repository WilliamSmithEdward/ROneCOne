const fs = require("node:fs/promises");
const path = require("node:path");
const { SpreadsheetFile, Workbook } = require("@oai/artifact-tool");

const root = path.resolve(__dirname, "..");
const outputDir = path.join(root, "demo", ".working");
const outputPath = path.join(outputDir, "ROneCOne_Delegates_Demo.xlsx");

const colors = {
  navy: "#152238",
  teal: "#00A6A6",
  orange: "#FF8C42",
  pale: "#EAF7F6",
  ink: "#1F2937",
  muted: "#64748B",
  line: "#D7E1EA",
  white: "#FFFFFF",
  green: "#15803D",
};

async function main() {
const workbook = Workbook.create();
const start = workbook.worksheets.add("Start Here");
const examples = workbook.worksheets.add("Examples");
const benchmarks = workbook.worksheets.add("Benchmarks");
const architecture = workbook.worksheets.add("Architecture");

function titleBand(sheet, title, subtitle, endColumn) {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endColumn}2`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endColumn}2`).format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 24 },
    verticalAlignment: "center",
  };
  sheet.getRange(`A3:${endColumn}3`).merge();
  sheet.getRange("A3").values = [[subtitle]];
  sheet.getRange(`A3:${endColumn}3`).format = {
    fill: colors.pale,
    font: { color: colors.ink, italic: true, size: 11 },
    verticalAlignment: "center",
  };
  sheet.getRange(`A1:${endColumn}3`).format.rowHeight = 24;
}

function section(range) {
  range.format = {
    fill: colors.teal,
    font: { bold: true, color: colors.white },
    borders: { preset: "outside", style: "thin", color: colors.teal },
  };
}

function tableHeader(range) {
  range.format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white },
    wrapText: true,
    borders: { preset: "all", style: "thin", color: colors.line },
  };
}

titleBand(
  start,
  "ROneCOne Delegates",
  "Build reusable pricing rules and send one update to several workbook features",
  "H",
);
start.getRange("A5:D5").merge();
start.getRange("A5").values = [["Run this demo"]];
section(start.getRange("A5:D5"));
start.getRange("A6:D9").values = [
  ["1", "Press Alt+F8", "Open Excel's Macro dialog", null],
  ["2", "Run RunROneCOneDemo", "Fills in every result for you", null],
  ["3", "Review Examples", "PASS confirms that each result worked", null],
  ["4", "Import ROneCOne.cls", "Add the entire runtime to your own workbook", null],
];
start.getRange("A6:D9").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
  wrapText: true,
};
start.getRange("A6:A9").format = {
  fill: colors.orange,
  font: { bold: true, color: colors.white },
  horizontalAlignment: "center",
};

start.getRange("F5:H5").merge();
start.getRange("F5").values = [["Live status"]];
section(start.getRange("F5:H5"));
start.getRange("F6:G10").values = [
  ["Examples passing", null],
  ["Examples total", null],
  ["Feature slice", "Universal delegates + expressions"],
  ["Runtime files", 1],
  ["Runtime dependencies", 0],
];
start.getRange("G6").formulas = [["=COUNTIF('Examples'!F6:F16,\"PASS\")"]];
start.getRange("G7").formulas = [["=COUNTA('Examples'!A6:A16)"]];
start.getRange("F6:G10").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
};
start.getRange("F6:F10").format.font = { bold: true, color: colors.ink };
start.getRange("G6:G7").format = {
  fill: colors.pale,
  font: { bold: true, color: colors.green, size: 14 },
  horizontalAlignment: "center",
};

start.getRange("A11:B11").values = [["Last demo run", "Value"]];
tableHeader(start.getRange("A11:B11"));
start.getRange("A12:A14").values = [["Timestamp"], ["Status"], ["Error detail"]];
start.getRange("A12:B14").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
  wrapText: true,
};
start.getRange("B12").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
start.getRange("B13").format = {
  fill: colors.pale,
  font: { bold: true, color: colors.green },
};
start.getRange("A16:H16").merge();
start.getRange("A16").values = [[
  "Your data stays local: no telemetry, network transmission, or runtime VBIDE access.",
]];
start.getRange("A16:H16").format = {
  fill: "#FFF4E8",
  font: { bold: true, color: "#9A4A00" },
  wrapText: true,
};
start.getRange("A1:H16").format.verticalAlignment = "center";
start.getRange("A:A").format.columnWidth = 16;
start.getRange("B:B").format.columnWidth = 22;
start.getRange("C:D").format.columnWidth = 28;
start.getRange("E:E").format.columnWidth = 3;
start.getRange("F:F").format.columnWidth = 22;
start.getRange("G:G").format.columnWidth = 28;
start.getRange("H:H").format.columnWidth = 10;
start.getRange("6:10").format.rowHeight = 28;
start.freezePanes.freezeRows(3);

titleBand(
  examples,
  "Live delegate examples",
  "Column E is written by RunROneCOneDemo; column F checks the result with worksheet formulas.",
  "F",
);
examples.getRange("A5:F5").values = [[
  "What it does",
  "C# equivalent (optional)",
  "ROneCOne VBA",
  "Expected",
  "Live result",
  "Status",
]];
tableHeader(examples.getRange("A5:F5"));
examples.getRange("A6:D16").values = [
  ["Apply a discount", "price => price * 0.9", "Set price = ROneCOne.Var(vbDouble)\nSet applyDiscount = price.Multiply(0.9).AsFunc", 90],
  ["Add shipping", "(amount, shipping) => amount + shipping", "Set orderTotal = amount.Add(shipping).AsFunc\norderTotal(100, 5)", 105],
  ["Check an approval range", "amount >= 100 && amount < 1000", "Set approvalRule = amount.AtLeast(100)\n    .AndAlso(amount.LessThan(1000)).AsFunc", true],
  ["Avoid unsafe work", "false && unsafeOperation", "ROneCOne.Value(False).AndAlso(...).AsFunc", false],
  ["Reuse an Excel function", "new Func<int,int,double>(Max)", "ROneCOne.Func(WorksheetFunction, \"Max\")\n    .Takes(vbLong, vbLong).Returns(vbDouble)", 7],
  ["Reuse workbook code", "new Func<int,int,int>(CalculateOrderTotal)", "ROneCOne.Func(\"DemoUsage.CalculateOrderTotal\")\n    .Takes(vbLong, vbLong).Returns(vbLong)", 105],
  ["Call with an input array", "calculateTotal.DynamicInvoke(args)", "calculateTotal.DynamicInvoke(Array(100, 5))", 105],
  ["Notify two features", "Delegate.Combine(dashboard, audit)", "Set notify = ROneCOne.Combine(updateDashboard, writeAudit)\nnotify.Execute \"Order 1042 approved\"", "Dashboard updated; audit written"],
  ["Update the original number", "increment(ref orderNumber)", "increment.Execute ROneCOne.RefLong(orderNumber)", 1042],
  ["Build a pricing pipeline", "discount.Then(addHandling)", "applyDiscount.PipeTo(addHandling)(100)", 95],
  ["Inspect the contract", "delegate.GetType()", "calculateTotal.Signature", "Func<Long, Long, Long>"],
];
examples.getRange("F6").formulas = [["=IF(E6=\"\",\"NOT RUN\",IF(E6=D6,\"PASS\",\"CHECK\"))"]];
examples.getRange("F6:F16").fillDown();
examples.getRange("A6:F16").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
  wrapText: true,
  verticalAlignment: "top",
};
examples.getRange("C6:C16").format = {
  fill: "#F8FAFC",
  font: { name: "Consolas", color: colors.ink, size: 10 },
  wrapText: true,
};
examples.getRange("F6:F16").conditionalFormats.add("containsText", {
  text: "PASS",
  format: { fill: "#DCFCE7", font: { bold: true, color: colors.green } },
});
examples.getRange("F6:F16").conditionalFormats.add("containsText", {
  text: "CHECK",
  format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
});
examples.getRange("A:A").format.columnWidth = 20;
examples.getRange("B:B").format.columnWidth = 25;
examples.getRange("C:C").format.columnWidth = 52;
examples.getRange("D:F").format.columnWidth = 15;
examples.getRange("6:16").format.rowHeight = 54;
examples.freezePanes.freezeRows(5);

titleBand(
  benchmarks,
  "Invocation benchmark",
  "Measured inside the same Excel process; no worker Excel instances are launched.",
  "F",
);
benchmarks.getRange("A5:E5").values = [[
  "Scenario",
  "Invocations",
  "Seconds",
  "Last result",
  "Invocations / second",
]];
tableHeader(benchmarks.getRange("A5:E5"));
benchmarks.getRange("A6").values = [["applyDiscount(amount) inferred AsFunc delegate"]];
benchmarks.getRange("E6").formulas = [["=IF(C6=0,0,B6/C6)"]];
benchmarks.getRange("A6:E6").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
};
benchmarks.getRange("C6").format.numberFormat = "0.000000";
benchmarks.getRange("E6").format.numberFormat = "#,##0";
benchmarks.getRange("A8:F8").merge();
benchmarks.getRange("A8").values = [[
  "This benchmark tracks expression-delegate overhead as a release-to-release performance baseline.",
]];
benchmarks.getRange("A8:F8").format = {
  fill: "#FFF4E8",
  font: { color: "#9A4A00" },
  wrapText: true,
};
benchmarks.getRange("A:A").format.columnWidth = 34;
benchmarks.getRange("B:E").format.columnWidth = 20;
benchmarks.getRange("F:F").format.columnWidth = 3;
benchmarks.getRange("8:8").format.rowHeight = 42;
benchmarks.freezePanes.freezeRows(5);

titleBand(
  architecture,
  "Why deployment stays simple",
  "Every capability is contained in ROneCOne.cls.",
  "F",
);
architecture.getRange("A5:F5").values = [[
  "Invariant",
  "Decision",
  "Why it matters",
  "Current status",
  "Runtime files",
  "Runtime installs",
]];
tableHeader(architecture.getRange("A5:F5"));
architecture.getRange("A6:F11").values = [
  ["Single-file core", "ROneCOne.cls", "One import operation", "ENFORCED", 1, 0],
  ["No runtime VBIDE", "Expression trees", "Works without trusted project access", "ENFORCED", 1, 0],
  ["Concise code", "Inferred Func + reusable adapters", "Less VBA ceremony", "ENFORCED", 1, 0],
  ["One Excel process", "Single-process execution contract", "No multi-instance parallelism", "ENFORCED", 1, 0],
  ["Privacy", "Local-only opt-in logs", "Never transmits workbook data", "ENFORCED", 1, 0],
  ["Workbook formats", ".xlsm / .xlsb / .xlam", "Normal VBA remains unchanged", "SUPPORTED CONTRACT", 1, 0],
];
architecture.getRange("A6:F11").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
  wrapText: true,
  verticalAlignment: "top",
};
architecture.getRange("D6:D11").format = {
  fill: colors.pale,
  font: { bold: true, color: colors.green },
};
architecture.getRange("A12:D12").values = [[
  "Milestone",
  "Capability",
  "Release status",
  "Depends on",
]];
tableHeader(architecture.getRange("A12:D12"));
architecture.getRange("A13:D18").values = [
  [1, "Universal delegates + expression lambdas", "AVAILABLE (v0.5.0)", "Tagged object kernel"],
  [2, "Runtime-generic List<T> + LINQ", "AVAILABLE (v0.2.0)", "Delegates"],
  [3, "Inferred Func + clear LINQ syntax", "AVAILABLE (v0.4.0)", "Delegates + collections"],
  [4, "Try / Catch / Finally", "AVAILABLE (v0.6.0)", "Delegates"],
  [5, "Typed events", "AVAILABLE (v0.6.0)", "Actions + multicast"],
  [6, "Tasks / async / await / cancellation", "AVAILABLE (v1.0.0)", "Exceptions + delegates"],
];
architecture.getRange("A13:D18").format = {
  borders: { preset: "all", style: "thin", color: colors.line },
  wrapText: true,
};
architecture.getRange("A:A").format.columnWidth = 22;
architecture.getRange("B:B").format.columnWidth = 40;
architecture.getRange("C:C").format.columnWidth = 24;
architecture.getRange("D:D").format.columnWidth = 32;
architecture.getRange("E:F").format.columnWidth = 16;
architecture.getRange("6:11").format.rowHeight = 42;
architecture.freezePanes.freezeRows(5);

await fs.mkdir(outputDir, { recursive: true });
const inspect = await workbook.inspect({
  kind: "table,formula",
  sheetId: "Examples",
  range: "A1:F16",
  maxChars: 4000,
  tableMaxRows: 17,
  tableMaxCols: 6,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
await fs.writeFile(
  path.join(outputDir, "delegates-inspect.ndjson"),
  `${inspect.ndjson}\n${errors.ndjson}`,
  "utf8",
);

for (const sheetName of ["Start Here", "Examples", "Benchmarks", "Architecture"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  const safeName = sheetName.toLowerCase().replaceAll(" ", "-");
  await fs.writeFile(
      path.join(outputDir, `delegates-${safeName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
