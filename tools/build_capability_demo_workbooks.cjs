const fs = require("node:fs/promises");
const path = require("node:path");
const { SpreadsheetFile, Workbook } = require("@oai/artifact-tool");

const root = path.resolve(__dirname, "..");
const outputDir = path.join(root, "demo", ".working");

const colors = {
  navy: "#152238",
  teal: "#00A6A6",
  orange: "#FF8C42",
  pale: "#EAF7F6",
  ink: "#1F2937",
  line: "#D7E1EA",
  white: "#FFFFFF",
  green: "#15803D",
};

const capabilities = [
  {
    key: "events",
    title: "ROneCOne Events",
    subtitle: "Strongly typed C#-style events over universal multicast Actions",
    macro: "RunROneCOneEventsDemo",
    feature: "Typed events",
    output: "ROneCOne_Events_Demo.xlsx",
    benchmark: "Event.Emit with one typed handler",
    benchmarkResult: "Accumulated value",
    examples: [
      [
        "Subscribe and Emit",
        "changed += first; changed += second",
        "Set changed = ROneCOne.EventOf(vbString)\n" +
          "    .Subscribe(first).Subscribe(second)\nchanged.Emit \"ready\"",
        "first:ready|second:ready|",
      ],
      ["Handler count", "invocationList.Length", "changed.HandlerCount", 2],
      ["Unsubscribe", "changed -= second", "changed.Unsubscribe second", true],
      ["Emit after removal", "changed?.Invoke(\"again\")", "changed.Emit \"again\"", "first:again|"],
    ],
  },
  {
    key: "exceptions",
    title: "ROneCOne Exceptions",
    subtitle: "Structured Try, Catch, filtered Catch, Finally, and rethrow semantics",
    macro: "RunROneCOneExceptionsDemo",
    feature: "Try / Catch / Finally",
    output: "ROneCOne_Exceptions_Demo.xlsx",
    benchmark: "Successful Try.Execute",
    benchmarkResult: "Trace characters",
    examples: [
      [
        "Catch and Finally",
        "try { work(); } catch (Exception e) { ... } finally { ... }",
        "Set attempt = ROneCOne.Try(work)\n" +
          "    .Catch(handler).Finally(cleanup)\nattempt.Execute",
        "caught:11|finally|",
      ],
      ["Filtered rethrow", "catch when error matches", ".Catch(errorNumber, handler)", 11],
      ["Finally on rethrow", "finally always runs", "ROneCOne.Try(work).Finally(cleanup)", "finally|"],
      ["Finally on success", "try { work(); } finally { cleanup(); }", "Set attempt = ROneCOne.Try(work).Finally(cleanup)\nattempt.Execute", "work|finally|"],
    ],
  },
];

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

async function buildCapability(config) {
  const workbook = Workbook.create();
  const start = workbook.worksheets.add("Start Here");
  const examples = workbook.worksheets.add("Examples");
  const benchmarks = workbook.worksheets.add("Benchmarks");
  const architecture = workbook.worksheets.add("Architecture");
  const lastRow = 5 + config.examples.length;

  titleBand(start, config.title, config.subtitle, "H");
  start.getRange("A5:D5").merge();
  start.getRange("A5").values = [["Run the living demo"]];
  section(start.getRange("A5:D5"));
  start.getRange("A6:D9").values = [
    ["1", "Press Alt+F8", "Open Excel's Macro dialog", null],
    ["2", `Run ${config.macro}`, "Executes every example and benchmark", null],
    ["3", "Review Examples", "Live results recalculate their PASS status", null],
    ["4", "Import ROneCOne.cls", "The runtime itself remains one file", null],
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
    ["Feature slice", config.feature],
    ["Runtime files", 1],
    ["Runtime dependencies", 0],
  ];
  start.getRange("G6").formulas = [[`=COUNTIF('Examples'!F6:F${lastRow},"PASS")`]];
  start.getRange("G7").formulas = [[`=COUNTA('Examples'!A6:A${lastRow})`]];
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
    "Privacy invariant: no telemetry, network transmission, or runtime VBIDE access.",
  ]];
  start.getRange("A16:H16").format = {
    fill: "#FFF4E8",
    font: { bold: true, color: "#9A4A00" },
    wrapText: true,
  };
  start.getRange("A:A").format.columnWidth = 16;
  start.getRange("B:B").format.columnWidth = 30;
  start.getRange("C:D").format.columnWidth = 26;
  start.getRange("F:F").format.columnWidth = 22;
  start.getRange("G:H").format.columnWidth = 19;
  start.getRange("6:10").format.rowHeight = 34;
  start.freezePanes.freezeRows(3);

  titleBand(examples, `Live ${config.feature.toLowerCase()} examples`, `Run ${config.macro}; column F validates every result.`, "F");
  examples.getRange("A5:F5").values = [[
    "Pattern", "C# idea", "ROneCOne VBA", "Expected", "Live result", "Status",
  ]];
  tableHeader(examples.getRange("A5:F5"));
  examples.getRange(`A6:D${lastRow}`).values = config.examples;
  examples.getRange("F6").formulas = [["=IF(E6=\"\",\"NOT RUN\",IF(E6=D6,\"PASS\",\"CHECK\"))"]];
  examples.getRange(`F6:F${lastRow}`).fillDown();
  examples.getRange(`A6:F${lastRow}`).format = {
    borders: { preset: "all", style: "thin", color: colors.line },
    wrapText: true,
    verticalAlignment: "top",
  };
  examples.getRange(`C6:C${lastRow}`).format = {
    fill: "#F8FAFC",
    font: { name: "Consolas", color: colors.ink, size: 10 },
    wrapText: true,
  };
  examples.getRange("A:A").format.columnWidth = 22;
  examples.getRange("B:B").format.columnWidth = 40;
  examples.getRange("C:C").format.columnWidth = 58;
  examples.getRange("D:F").format.columnWidth = 18;
  examples.getRange(`6:${lastRow}`).format.rowHeight = 62;
  examples.freezePanes.freezeRows(5);

  titleBand(benchmarks, `${config.feature} benchmark`, "Measured in the same Excel process.", "F");
  benchmarks.getRange("A5:D5").values = [[
    "Scenario", "Iterations", "Seconds", config.benchmarkResult,
  ]];
  tableHeader(benchmarks.getRange("A5:D5"));
  benchmarks.getRange("A6").values = [[config.benchmark]];
  benchmarks.getRange("A6:D6").format = {
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  benchmarks.getRange("C6").format.numberFormat = "0.000000";
  benchmarks.getRange("A:A").format.columnWidth = 38;
  benchmarks.getRange("B:D").format.columnWidth = 20;
  benchmarks.freezePanes.freezeRows(5);

  titleBand(architecture, "One-file architecture", "Every value is a tagged role inside ROneCOne.cls.", "F");
  architecture.getRange("A5:F5").values = [[
    "Invariant", "Decision", "Behavior", "Status", "Runtime files", "Dependencies",
  ]];
  tableHeader(architecture.getRange("A5:F5"));
  architecture.getRange("A6:F10").values = [
    ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
    ["Strong typing", "Runtime signatures", "Fail before invocation", "ENFORCED", 1, 0],
    ["One process", "In-process dispatch", "Never launches Excel", "ENFORCED", 1, 0],
    ["No runtime VBIDE", "Tagged object kernel", "Normal macro security", "ENFORCED", 1, 0],
    ["Privacy", "No transmission", "Workbook data remains local", "ENFORCED", 1, 0],
  ];
  architecture.getRange("A6:F10").format = {
    borders: { preset: "all", style: "thin", color: colors.line },
    wrapText: true,
  };
  architecture.getRange("D6:D10").format = {
    fill: colors.pale,
    font: { bold: true, color: colors.green },
  };
  architecture.getRange("A:A").format.columnWidth = 22;
  architecture.getRange("B:B").format.columnWidth = 32;
  architecture.getRange("C:C").format.columnWidth = 32;
  architecture.getRange("D:F").format.columnWidth = 18;

  const inspect = await workbook.inspect({
    kind: "workbook,sheet,table,formula",
    maxChars: 5000,
    tableMaxRows: 12,
    tableMaxCols: 8,
  });
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: `${config.key} formula error scan`,
  });
  await fs.writeFile(
    path.join(outputDir, `${config.key}-inspect.ndjson`),
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
      path.join(outputDir, `${config.key}-${safeName}.png`),
      new Uint8Array(await preview.arrayBuffer()),
    );
  }
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  const outputPath = path.join(outputDir, config.output);
  await xlsx.save(outputPath);
  console.log(outputPath);
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });
  for (const capability of capabilities) {
    await buildCapability(capability);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
