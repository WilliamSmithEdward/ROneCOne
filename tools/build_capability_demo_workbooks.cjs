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
    subtitle: "Send one checked order update to several workbook features",
    macro: "RunROneCOneEventsDemo",
    feature: "Typed events",
    output: "ROneCOne_Events_Demo.xlsx",
    benchmark: "Event.Emit with one typed handler",
    benchmarkResult: "Accumulated value",
    examples: [
      [
        "Notify two features",
        "orderChanged += dashboard; orderChanged += audit",
        "Set orderChanged = ROneCOne.EventOf(vbString)\n" +
          "    .Subscribe(dashboard).Subscribe(audit)\n" +
          "orderChanged.Emit \"Order 1042 shipped\"",
        "Dashboard updated; audit written",
      ],
      ["Count listeners", "invocationList.Length", "orderChanged.HandlerCount", 2],
      ["Stop audit updates", "orderChanged -= audit", "orderChanged.Unsubscribe audit", true],
      [
        "Notify remaining feature",
        "orderChanged?.Invoke(message)",
        "orderChanged.Emit \"Order 1043 delayed\"",
        "Dashboard updated",
      ],
    ],
  },
  {
    key: "exceptions",
    title: "ROneCOne Exceptions",
    subtitle: "Protect a sales import and always close the source file",
    macro: "RunROneCOneExceptionsDemo",
    feature: "Try / Catch / Finally",
    output: "ROneCOne_Exceptions_Demo.xlsx",
    benchmark: "Successful import wrapped in Try.Execute",
    benchmarkResult: "Trace characters",
    examples: [
      [
        "Protect a sales import",
        "try { ImportSales(); } catch (InvalidAmount) { ... } finally { ... }",
        "Set attempt = ROneCOne.Try(importSales)\n" +
          "    .Catch(INVALID_AMOUNT_ERROR, skipBadRow)\n" +
          "    .Finally(closeFile)\nattempt.Execute",
        "3 rows imported; file closed",
      ],
      ["Keep recovery ready", "catch (InvalidAmount)", ".Catch(INVALID_AMOUNT_ERROR, skipBadRow)", true],
      ["Confirm a clean import", "no exception", "The Catch handler was not needed", "No import error"],
      ["Close after success", "try { ImportSales(); } finally { CloseFile(); }", "Set attempt = ROneCOne.Try(validImport).Finally(closeFile)\nattempt.Execute", "3 rows imported; file closed"],
    ],
  },
  {
    key: "tasks",
    title: "ROneCOne Tasks",
    subtitle: "Coordinate calculations, cancellation, and progress on Excel's thread",
    macro: "RunROneCOneTasksDemo",
    feature: "Tasks + async",
    output: "ROneCOne_Tasks_Demo.xlsx",
    benchmark: "Cooperative Task.Run startup + Await",
    benchmarkResult: "Last result",
    examples: [
      ["Run two calculations", "await Task.WhenAll(forecast, reorder)", "Set results = ROneCOne.Task.WhenAll(forecastTask, reorderTask).Await", "135000 | 152"],
      ["Build the next step", "allWork.ContinueWith(BuildSummary)", "allWork.ContinueWith(buildSummary).Await", "Forecast 135000; reorder point 152"],
      ["Keep workbook work safe", "run on the UI thread", "ROneCOne.Task.Run(countOpenOrders).Await", 3],
      ["Pause without another Excel", "await Task.Delay(5)", "ignored = ROneCOne.Task.Delay(5&).Await", true],
      ["Cancel safely", "cancelSource.Cancel()", "source.Cancel: source.Token.IsCancellationRequested", true],
      ["Show progress", "progress.Report(7)", "ROneCOne.ProgressOf(vbLong, handler).Report 7&", 7],
      ["Finish from a callback", "source.SetResult(99)", "completion.SetResult 99&: completion.Task.Await", 99],
      ["Limit waiting time", "await task.WaitAsync(timeout)", "task.WaitAsync(100&).Await", true],
      ["Let Excel breathe", "await Task.Yield()", "ignored = ROneCOne.Task.YieldOnce.Await", true],
    ],
  },
  {
    key: "data",
    title: "ROneCOne Data + Providers",
    subtitle: "Build validated tables, query them, and load local Excel data",
    macro: "RunROneCOneDataDemo",
    feature: "Data + providers",
    output: "ROneCOne_Data_Demo.xlsx",
    benchmark: "Build 1,000 typed rows and query them",
    benchmarkResult: "Selected rows",
    examples: [
      ["Add a validated row", "DataColumn + Rows.Add", "table.Column(\"Id\", vbLong).AutoNumber(100&, 10&).AsPrimaryKey\nSet row = table.Row(\"Ada\", 90&, ROneCOne.DBNull).Add", 100],
      ["Show the top score", "view.Sort + RowFilter", "DataView(table).WithFilter(...).WithSort(\"Score\", True)", "Grace"],
      ["Connect customers to orders", "parent.GetChildRows(...) ", "parentRow.GetChildRows(\"CustomerOrders\").Count", 1],
      ["Find unsaved changes", "table.GetChanges()", "table.GetChanges.Rows.Count", 1],
      ["Load an Excel table", "adapter.Fill(table)", "ROneCOne.DbDataAdapter(command).Fill(filled)", 2],
      ["Keep source ordering", "reader.GetString(0)", "filled.Rows.Item(0).Item(\"Name\")", "Grace"],
      ["Await a record count", "await command.ExecuteScalarAsync()", "command.ExecuteScalarAsync.Await", 2],
      ["Store a blank database value", "DBNull.Value", "ROneCOne.DBNull", true],
      ["See how waiting works", "provider capability inspection", "connection.AsyncMode", "Cooperative"],
      ["Confirm safe single-thread use", "provider capability inspection", "connection.SupportsNativeAsync", false],
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
  start.getRange("A5").values = [["Run this demo"]];
  section(start.getRange("A5:D5"));
  start.getRange("A6:D9").values = [
    ["1", "Press Alt+F8", "Open Excel's Macro dialog", null],
    ["2", `Run ${config.macro}`, "Fills in every result for you", null],
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
    "Your data stays local: no telemetry, network transmission, or runtime VBIDE access.",
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
    "What it does", "C# equivalent (optional)", "ROneCOne VBA", "Expected", "Live result", "Status",
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

  titleBand(architecture, "Why deployment stays simple", "Every capability is contained in ROneCOne.cls.", "F");
  architecture.getRange("A5:F5").values = [[
    "Invariant", "Decision", "Behavior", "Status", "Runtime files", "Dependencies",
  ]];
  tableHeader(architecture.getRange("A5:F5"));
  architecture.getRange("A6:F10").values = [
    ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
    ["Checked values", "Runtime signatures", "Reject mistakes before work starts", "ENFORCED", 1, 0],
    ["One process", "Cooperative scheduler", "Never launches another Excel", "ENFORCED", 1, 0],
    ["No runtime VBIDE", "One internal object model", "Normal macro security", "ENFORCED", 1, 0],
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
