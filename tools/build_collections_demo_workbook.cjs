const fs = require("node:fs/promises");
const path = require("node:path");
const { SpreadsheetFile, Workbook } = require("@oai/artifact-tool");

const root = path.resolve(__dirname, "..");
const outputDir = path.join(root, "demo", ".working");
const outputPath = path.join(outputDir, "ROneCOne_Collections_Demo.xlsx");

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

async function main() {
  const workbook = Workbook.create();
  const start = workbook.worksheets.add("Start Here");
  const examples = workbook.worksheets.add("Examples");
  const userClassLinq = workbook.worksheets.add("User Class LINQ");
  const benchmarks = workbook.worksheets.add("Benchmarks");
  const architecture = workbook.worksheets.add("Architecture");

  titleBand(
    start,
    "ROneCOne Collections + LINQ",
    "Filter, sort, and summarize ordinary VBA objects without writing loop after loop",
    "H",
  );
  start.getRange("A5:D5").merge();
  start.getRange("A5").values = [["Run this demo"]];
  section(start.getRange("A5:D5"));
  start.getRange("A6:D9").values = [
    ["1", "Press Alt+F8", "Open Excel's Macro dialog", null],
    ["2", "Run RunROneCOneCollectionsDemo", "Fills in every result for you", null],
    ["3", "Review both example sheets", "PASS confirms that each result worked", null],
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
    ["Feature slice", "List<T> + LINQ"],
    ["Runtime files", 1],
    ["Runtime dependencies", 0],
  ];
  start.getRange("G6").formulas = [[
    "=COUNTIF('Examples'!F6:F13,\"PASS\")+" +
      "COUNTIF('User Class LINQ'!F6:F22,\"PASS\")",
  ]];
  start.getRange("G7").formulas = [[
    "=COUNTA('Examples'!A6:A13)+COUNTA('User Class LINQ'!A6:A22)",
  ]];
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
  start.getRange("B:B").format.columnWidth = 25;
  start.getRange("C:D").format.columnWidth = 28;
  start.getRange("E:E").format.columnWidth = 3;
  start.getRange("F:F").format.columnWidth = 22;
  start.getRange("G:H").format.columnWidth = 20;
  start.getRange("6:10").format.rowHeight = 28;
  start.freezePanes.freezeRows(3);

  titleBand(
    examples,
    "Live List<T> and LINQ examples",
    "Column E is written by the macro; column F verifies it with worksheet formulas.",
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
  examples.getRange("A6:D13").values = [
    ["Create a checked number list", "new List<long> { 5, 10, 15 }", "ROneCOne.ListOf(vbLong, 5, 10, 15)", "List<Long>"],
    ["Reject the wrong data type", "Compile-time element type", "values.Add \"not a Long\"", true],
    ["Use existing Customer objects", "new List<DemoCustomer> { ada, grace }", "ROneCOne.ListFrom(ada, grace)", "List<DemoCustomer>; second customer: Grace"],
    ["See newly added matches", "query observes later mutation", "Set element = values.Element\nSet query = values.Where(element.AtLeast(10))\nvalues.Add 30", "2 matches; last: 30"],
    ["Filter and rank values", "Where.Select.OrderBy.Take", ".Where(...).Map(...).OrderDescending.Take(2)", "Top results: 60, 40"],
    ["Clean and reshape values", "Distinct.Prepend.Append.Reverse.Skip", "values.Distinct.Prepend(1).Append(4).Reverse.Skip(1)", "Sequence: 3, 2, 1"],
    ["Summarize numbers", "Sum/Average/Min/Max", "Range(1, 5).Sum / Average / Min / Max", "Sum 15; average 3; min 1; max 5"],
    ["Run work for every item", "values.ForEach(action)", "values.ForEach ROneCOne.Action(...)", 10],
  ];
  examples.getRange("F6").formulas = [
    ["=IF(E6=\"\",\"NOT RUN\",IF(E6=D6,\"PASS\",\"CHECK\"))"],
  ];
  examples.getRange("F6:F13").fillDown();
  examples.getRange("A6:F13").format = {
    borders: { preset: "all", style: "thin", color: colors.line },
    wrapText: true,
    verticalAlignment: "top",
  };
  examples.getRange("C6:C13").format = {
    fill: "#F8FAFC",
    font: { name: "Consolas", color: colors.ink, size: 10 },
    wrapText: true,
  };
  examples.getRange("F6:F13").conditionalFormats.add("containsText", {
    text: "PASS",
    format: { fill: "#DCFCE7", font: { bold: true, color: colors.green } },
  });
  examples.getRange("F6:F13").conditionalFormats.add("containsText", {
    text: "CHECK",
    format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
  });
  examples.getRange("A:A").format.columnWidth = 19;
  examples.getRange("B:B").format.columnWidth = 27;
  examples.getRange("C:C").format.columnWidth = 54;
  examples.getRange("D:D").format.columnWidth = 25;
  examples.getRange("E:E").format.columnWidth = 26;
  examples.getRange("F:F").format.columnWidth = 16;
  examples.getRange("6:13").format.rowHeight = 54;
  examples.freezePanes.freezeRows(5);

  titleBand(
    userClassLinq,
    "User-defined class LINQ",
    "Query ordinary Customer objects with readable property names and no adapter class.",
    "F",
  );
  userClassLinq.getRange("A5:F5").values = [[
    "What it does",
    "C# equivalent (optional)",
    "ROneCOne VBA",
    "Expected",
    "Live result",
    "Status",
  ]];
  tableHeader(userClassLinq.getRange("A5:F5"));
  userClassLinq.getRange("A6:D22").values = [
    [
      "Use Customer objects directly",
      "List<DemoCustomer>",
      "Set customers = ROneCOne.ListFrom(ada, grace, katherine)",
      "List<DemoCustomer> with 4 customers",
    ],
    [
      "Include customers added later",
      "customers.Where(c => c.Age >= 40)",
      'Set experienced = customers.Where("Age").AtLeast(40)\n' +
        "customers.Add margaret",
      "3 customers; newest match: Margaret",
    ],
    [
      "Get and alphabetize names",
      ".Select(c => c.Name).OrderBy(name => name)",
      'Set names = experienced.Map("CustomerName", vbString).Order.ToList',
      "Grace, Katherine, Margaret",
    ],
    [
      "Sort by city, then age",
      ".OrderBy(c => c.City).ThenByDescending(c => c.Age).First()",
      'Set first = customers.OrderBy("City").ThenByDescending("Age").First',
      "First: Margaret, age 45",
    ],
    [
      "Ask yes-or-no questions",
      ".Any(city) / .All(age)",
      'customers.Exists(customers.Condition("City").EqualTo("London"))',
      "London exists: True; all age 40+: False",
    ],
    [
      "Calculate the average age",
      ".Where(age).Average(c => c.Age)",
      'averageAge = experienced.Average("Age")',
      44.7,
    ],
    [
      "Reuse an existing VBA rule",
      "customers.Where(IsExperienced)",
      'customers.WhereMethod("CollectionsDemoUsage.IsExperiencedCustomer")',
      "3 matches; Func<DemoCustomer, Boolean>",
    ],
    [
      "Search within names",
      "StartsWith / Contains",
      '.Where("CustomerName").StartsWith("G")\n' +
        '.Where("CustomerName").Contains("ther")',
      "Starts with G: 1; contains ther: 1",
    ],
    [
      "Count unique cities",
      ".DistinctBy(c => c.City)",
      'customers.DistinctBy("City")',
      3,
    ],
    [
      "Use concise !Age syntax",
      "c => c.Age >= 40",
      'With customers\n  Set q = .Where(!Age.AtLeast(40))\nEnd With',
      "Grace, Katherine, Margaret",
    ],
    [
      "Handle a missing manager",
      "c.Manager?.Age >= 40",
      '.Where("Manager?.Age").AtLeast(40)',
      "No manager: 1; manager age 40+: 2",
    ],
    [
      "Filter by an allowed-city list",
      "allowedCities.Contains(c.City)",
      '.Where("City").IsIn(allowedCities)\n' +
        ".Where(allowedCities.Contains(customers!City))",
      "Allowed city matches: 2; expression matches: 2",
    ],
    [
      "Match text without case",
      "StringComparison.OrdinalIgnoreCase",
      '.Where("City").EqualToIgnoreCase("LONDON")\n' +
        '.Where("CustomerName").ContainsIgnoreCase("THER")',
      "City matches: 1; name contains: 1",
    ],
    [
      "Count or find exact matches",
      "Count(predicate) / Single() / None()",
      'customers.Count(agePredicate)\n' +
        'customers.SingleItem(customers.Match("CustomerName", "Grace"))',
      "Count: 3; single: Grace; none age 100: True",
    ],
    [
      "Check each manager's reports",
      "Reports.Any / All / !Any",
      'customers.WhereAny("Reports", reportPredicate)\n' +
        'customers.WhereAll("Reports", allPredicate)\n' +
        'customers.WhereNone("Reports", reportPredicate)',
      "Any: 2; all: 4; none: 2",
    ],
    [
      "Customize text matching",
      "IEqualityComparer<T> / IComparer<T>",
      "strings.Distinct(equalityComparer)\n" +
        "strings.Contains(\"ada\", equalityComparer)\n" +
        "strings.Order(comparer)",
      "Distinct: 2; contains Ada: True; first: Ada",
    ],
    [
      "Combine yes-or-no rules",
      "predicate && predicate / predicate || predicate",
      "customers.Where(agePredicate.Both(cityPredicate))\n" +
        "customers.Where(london.Either(cleveland))",
      "Both: 2; either: 2",
    ],
  ];
  userClassLinq.getRange("F6").formulas = [[
    "=IF(E6=\"\",\"NOT RUN\",IF(E6=D6,\"PASS\",\"CHECK\"))",
  ]];
  userClassLinq.getRange("F6:F22").fillDown();
  userClassLinq.getRange("A6:F22").format = {
    borders: { preset: "all", style: "thin", color: colors.line },
    wrapText: true,
    verticalAlignment: "top",
  };
  userClassLinq.getRange("C6:C22").format = {
    fill: "#F8FAFC",
    font: { name: "Consolas", color: colors.ink, size: 10 },
    wrapText: true,
  };
  userClassLinq.getRange("D11:E11").format.numberFormat = "0.0";
  userClassLinq.getRange("F6:F22").conditionalFormats.add("containsText", {
    text: "PASS",
    format: { fill: "#DCFCE7", font: { bold: true, color: colors.green } },
  });
  userClassLinq.getRange("F6:F22").conditionalFormats.add("containsText", {
    text: "CHECK",
    format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
  });
  userClassLinq.getRange("A:A").format.columnWidth = 23;
  userClassLinq.getRange("B:B").format.columnWidth = 36;
  userClassLinq.getRange("C:C").format.columnWidth = 58;
  userClassLinq.getRange("D:D").format.columnWidth = 31;
  userClassLinq.getRange("E:E").format.columnWidth = 25;
  userClassLinq.getRange("F:F").format.columnWidth = 16;
  userClassLinq.getRange("6:22").format.rowHeight = 62;
  userClassLinq.freezePanes.freezeRows(5);

  titleBand(
    benchmarks,
    "Typed query benchmark",
    "10,000 elements in one Excel process; no worker Excel instances are launched.",
    "F",
  );
  benchmarks.getRange("A5:E5").values = [[
    "Scenario",
    "Source elements",
    "Seconds",
    "Result elements",
    "Elements / second",
  ]];
  tableHeader(benchmarks.getRange("A5:E5"));
  benchmarks.getRange("A6:A10").values = [
    ["Range.Where.ToList"],
    ["Repeated object member Where"],
    ["OrderBy.ThenByDescending.ToList"],
    ["Dictionary Add + Item (10K)"],
    ["Dictionary Item (100K)"],
  ];
  benchmarks.getRange("E6").formulas = [["=IF(C6=0,0,B6/C6)"]];
  benchmarks.getRange("E6:E10").fillDown();
  benchmarks.getRange("A6:E10").format = {
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  benchmarks.getRange("B6:B10").format.numberFormat = "#,##0";
  benchmarks.getRange("C6:C10").format.numberFormat = "0.000000";
  benchmarks.getRange("D6:E10").format.numberFormat = "#,##0";
  benchmarks.getRange("A12:F12").merge();
  benchmarks.getRange("A12").values = [[
    "Release gates: filtering, composite ordering, and indexed dictionary workloads stay bounded in one Excel process.",
  ]];
  benchmarks.getRange("A12:F12").format = {
    fill: "#FFF4E8",
    font: { color: "#9A4A00" },
    wrapText: true,
  };
  benchmarks.getRange("A:A").format.columnWidth = 34;
  benchmarks.getRange("B:E").format.columnWidth = 20;
  benchmarks.freezePanes.freezeRows(5);

  titleBand(
    architecture,
    "How collections stay safe and fast",
    "One ROneCOne class checks values, builds queries, and keeps lookups responsive.",
    "F",
  );
  architecture.getRange("A5:F5").values = [[
    "Invariant",
    "Decision",
    "Behavior",
    "Status",
    "Runtime files",
    "Dependencies",
  ]];
  tableHeader(architecture.getRange("A5:F5"));
  architecture.getRange("A6:F15").values = [
    ["One element type", "VarType or exact class name", "Reject before mutation", "ENFORCED", 1, 0],
    ["Existing VBA classes", "ListFrom or prototype token", "Preserve exact class identity", "ENFORCED", 1, 0],
    ["Up-to-date queries", "Immutable query nodes", "Evaluate when results are requested", "ENFORCED", 1, 0],
    ["Concise syntax", "Contextual members + native bang", "Full API remains available", "ENFORCED", 1, 0],
    ["Combined filter rules", "Composable immutable nodes", "Lists, missing values, and child records", "ENFORCED", 1, 0],
    ["Predictable sorting", "OrderBy + ThenBy chain", "Cached keys and O(n log n) merge sort", "ENFORCED", 1, 0],
    ["Fast key lookup", "Open addressing + direct value slots", "Average O(1) keyed lookup", "ENFORCED", 1, 0],
    ["Enumeration", "Version-cached materialization", "Nested For Each works", "ENFORCED", 1, 0],
    ["One process", "In-process execution", "Never launches Excel", "ENFORCED", 1, 0],
    ["Privacy", "No transmission", "Workbook data stays local", "ENFORCED", 1, 0],
  ];
  architecture.getRange("A6:F15").format = {
    borders: { preset: "all", style: "thin", color: colors.line },
    wrapText: true,
    verticalAlignment: "top",
  };
  architecture.getRange("D6:D15").format = {
    fill: colors.pale,
    font: { bold: true, color: colors.green },
  };
  architecture.getRange("A:A").format.columnWidth = 22;
  architecture.getRange("B:B").format.columnWidth = 32;
  architecture.getRange("C:C").format.columnWidth = 30;
  architecture.getRange("D:D").format.columnWidth = 18;
  architecture.getRange("E:F").format.columnWidth = 16;
  architecture.getRange("6:14").format.rowHeight = 40;
  architecture.freezePanes.freezeRows(5);

  await fs.mkdir(outputDir, { recursive: true });
  const inspect = await workbook.inspect({
    kind: "workbook,sheet,table,formula",
    maxChars: 7000,
    tableMaxRows: 15,
    tableMaxCols: 8,
  });
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
  });
  await fs.writeFile(
    path.join(outputDir, "collections-inspect.ndjson"),
    `${inspect.ndjson}\n${errors.ndjson}`,
    "utf8",
  );

  for (const sheetName of [
    "Start Here",
    "Examples",
    "User Class LINQ",
    "Benchmarks",
    "Architecture",
  ]) {
    const preview = await workbook.render({
      sheetName,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    const safeName = sheetName.toLowerCase().replaceAll(" ", "-");
    await fs.writeFile(
      path.join(outputDir, `collections-${safeName}.png`),
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
