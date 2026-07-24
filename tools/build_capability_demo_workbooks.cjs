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
      ["Pause without another Excel", "await Task.Delay(5)", "ignored = ROneCOne.Task.Delay(5).Await", true],
      ["Cancel safely", "cancelSource.Cancel()", "source.Cancel: source.Token.IsCancellationRequested", true],
      ["Show progress", "progress.Report(7)", "ROneCOne.ProgressOf(vbLong, handler).Report 7", 7],
      ["Finish from a callback", "source.SetResult(99)", "completion.SetResult 99: completion.Task.Await", 99],
      ["Limit waiting time", "await task.WaitAsync(timeout)", "task.WaitAsync(100).Await", true],
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
      ["Add a validated row", "DataColumn + Rows.Add", "table.Column(\"Id\", vbLong).AutoNumber(100, 10).AsPrimaryKey\nSet row = table.Row(\"Ada\", 90, ROneCOne.DBNull).Add", 100],
      ["Show the top score", "view.Sort + RowFilter", "DataView(table).WithFilter(...).WithSort(\"Score\", True)", "Grace"],
      ["Connect customers to orders", "parent.GetChildRows(...) ", "parentRow.GetChildRows(\"CustomerOrders\").Count", 1],
      ["Find unsaved changes", "table.GetChanges()", "table.GetChanges.Rows.Count", 1],
      ["Load an Excel table", "adapter.Fill(table)", "ROneCOne.DbDataAdapter(command).Fill(filled)", 2],
      ["Keep source ordering", "reader.GetString(0)", "filled.Rows.Item(0).Item(\"Name\")", "Grace"],
      ["Await a record count", "await command.ExecuteScalarAsync()", "command.ExecuteScalarAsync.Await", 2],
      ["Store a blank database value", "DBNull.Value", "ROneCOne.DBNull", true],
      ["See how queries wait", "provider capability inspection", "connection.AsyncMode", "Native"],
      ["Confirm the connection is open", "provider state inspection", "connection.State", "Open"],
    ],
  },
  {
    key: "http",
    title: "ROneCOne HTTP",
    subtitle: "Download web data with awaitable requests that keep Excel responsive",
    macro: "RunROneCOneHttpDemo",
    feature: "HTTP + async",
    output: "ROneCOne_Http_Demo.xlsx",
    benchmark: "Three downloads: one after another vs overlapped",
    benchmarkResult: "Sequential seconds",
    notice:
      "This demo requests data from https://pokeapi.co over the internet. " +
      "Only the URLs shown are contacted; the runtime itself transmits nothing.",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["In-process transport", "WinHttp.WinHttpRequest.5.1", "No installs, add-ins, or references", "ENFORCED", 1, 0],
      ["One VBA thread", "Cooperative Await", "Transfers overlap inside WinHTTP, never in VBA", "ENFORCED", 1, 0],
      ["Explicit network use", "Only URLs you request", "The runtime never phones home", "ENFORCED", 1, 0],
      ["Checked failures", "HttpRequestException", "Non-success raises a typed, catchable error", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Download one resource",
        "var response = await client.GetAsync(url);",
        "Set response = client.GetAsync(\"pokemon/pikachu\").Await",
        "200 OK",
      ],
      [
        "Read the body as text",
        "await client.GetStringAsync(url)",
        "json = client.GetStringAsync(\"pokemon/ditto\").Await",
        true,
      ],
      [
        "Check before you trust",
        "response.EnsureSuccessStatusCode();",
        "response.EnsureSuccessStatusCode",
        true,
      ],
      [
        "Read one header",
        "response.Content.Headers.ContentType",
        "response.Header(\"Content-Type\")",
        "application/json; charset=utf-8",
      ],
      [
        "Overlap three downloads",
        "await Task.WhenAll(first, second, third)",
        "Set replies = ROneCOne.Task.WhenAll(first, second, third).Await",
        "bulbasaur, charmander, squirtle ready",
      ],
      [
        "Handle a missing resource",
        "response.StatusCode == HttpStatusCode.NotFound",
        "client.GetAsync(\"pokemon/missingno\").Await.StatusCode",
        404,
      ],
      [
        "Catch a failed download",
        "try { await client.GetStringAsync(url); }\ncatch (HttpRequestException) { ... }",
        "On Error Resume Next\ntask.Await\nIf Err.Number = ROneCOne.HttpRequestError Then ...",
        "skipped a missing resource",
      ],
      [
        "Cancel a request",
        "cancelSource.Cancel();",
        "source.Cancel\nSet task = client.GetAsync(\"pokemon/eevee\", source.Token)",
        true,
      ],
      [
        "Reuse a base address",
        "client.BaseAddress = new Uri(baseUrl);",
        "client.BaseAddress = \"https://pokeapi.co/api/v2/\"\nSet berry = client.GetAsync(\"berry/1\").Await",
        true,
      ],
      [
        "Read one JSON field",
        "JsonNode.Parse(json)[\"name\"]",
        "Set tree = ROneCOne.Json.Deserialize(pikachuJson)\ntree.Item(\"name\")",
        "pikachu",
      ],
      [
        "Turn a response into a table",
        "JsonSerializer.Deserialize<List<Ability>>(json)",
        "Set abilities = ROneCOne.Json.DeserializeTable( _\n    pikachuJson, \"Abilities\", \"$.abilities\")",
        true,
      ],
      [
        "Download a body to a file",
        "await client.GetStreamAsync(url).CopyToAsync(file)",
        "client.DownloadFileAsync(\"pokemon/pikachu\", downloadPath).Await\nROneCOne.File.Exists(downloadPath)",
        true,
      ],
    ],
  },
  {
    key: "json",
    title: "ROneCOne JSON",
    subtitle: "Parse, build, and bind JSON without leaving the workbook",
    macro: "RunROneCOneJsonDemo",
    feature: "JSON",
    output: "ROneCOne_Json_Demo.xlsx",
    benchmark: "Round-trip a 1,000-order table",
    benchmarkResult: "Round-tripped rows",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["Strict parsing", "RFC 8259", "Malformed JSON raises a typed, positioned error", "ENFORCED", 1, 0],
      ["Native model", "Dictionaries + lists", "A parsed tree is queryable immediately", "ENFORCED", 1, 0],
      ["Locale-proof numbers", "Invariant separator", "Serialized numbers always use a dot", "ENFORCED", 1, 0],
      ["Explicit binding", "Factories + property names", "No reflection, no Activator, no magic", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Parse and read one member",
        "JsonNode.Parse(json)[\"customer\"]",
        "Set tree = ROneCOne.Json.Deserialize(body)\ntree.Item(\"customer\")",
        "Ada",
      ],
      [
        "Walk into nested data",
        "node[\"address\"][\"city\"]",
        "tree.Item(\"address\").Item(\"city\")",
        "London",
      ],
      [
        "Index into an array",
        "node[\"tags\"][1]",
        "tree.Item(\"tags\").Item(1)",
        "priority",
      ],
      [
        "Serialize a collection",
        "JsonSerializer.Serialize(new[] { 1, 2, 3 })",
        "ROneCOne.ListOf(vbLong, 1, 2, 3).ToJson",
        "[1,2,3]",
      ],
      [
        "Round-trip indented output",
        "new JsonSerializerOptions { WriteIndented = true }",
        "pretty = ROneCOne.Json.Serialize(tree, True)\nSet again = ROneCOne.Json.Deserialize(pretty)",
        true,
      ],
      [
        "Land an array in a DataTable",
        "JsonSerializer.Deserialize<List<Order>>(json)",
        "Set orders = ROneCOne.Json.DeserializeTable(ordersJson, \"Orders\")",
        3,
      ],
      [
        "Read a dotted nested column",
        "flatten with a custom converter",
        "orders.Rows.Item(0).Item(\"customer.city\")",
        "London",
      ],
      [
        "Address an envelope path",
        "root.GetProperty(\"data\").GetProperty(\"items\")",
        "ROneCOne.Json.DeserializeTable(envelope, \"Items\", \"$.data.items\")",
        2,
      ],
      [
        "Bind onto your class",
        "JsonSerializer.Deserialize<Customer>(json)",
        "ROneCOne.Json.DeserializeInto bindBody, customer\ncustomer.CustomerName",
        "Grace",
      ],
      [
        "Build objects through a factory",
        "JsonSerializer.Deserialize<List<Customer>>(json)",
        "Set factory = ROneCOne.Func(\"JsonDemoUsage.NewDemoCustomer\") _\n    .Takes().Returns(vbObject)\nSet people = ROneCOne.Json.DeserializeObjects(peopleJson, factory)",
        "Bo",
      ],
      [
        "Turn objects back into a table",
        "custom mapping code",
        "ROneCOne.DataTableFromObjects(people, _\n    Array(\"CustomerName\", \"Age\", \"City\"))",
        2,
      ],
      [
        "Reject malformed JSON",
        "JsonException with position info",
        "On Error Resume Next\nROneCOne.Json.Deserialize \"{\"\"a\"\":1,}\"\nIf Err.Number = ROneCOne.JsonError Then ...",
        "trailing comma rejected",
      ],
    ],
  },
  {
    key: "files",
    title: "ROneCOne Files + CSV",
    subtitle: "Read, write, and round-trip real files without leaving the workbook",
    macro: "RunROneCOneFilesDemo",
    feature: "Files + CSV",
    output: "ROneCOne_Files_Demo.xlsx",
    benchmark: "Round-trip 1,000 rows via a CSV file",
    benchmarkResult: "Round-tripped rows",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["Real UTF-8", "ADODB.Stream", "No byte-order mark on write, mark-driven reads", "ENFORCED", 1, 0],
      ["System.IO semantics", "File / Directory / Path", "C# missing and existing rules, typed IOError", "ENFORCED", 1, 0],
      ["RFC 4180 CSV", "Csv + table.ToCsv", "Strict quoting, deterministic column types", "ENFORCED", 1, 0],
      ["Self-contained", "One demo folder", "Everything written here is deleted at the end", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Write and read UTF-8 text",
        "File.WriteAllText(path, text)",
        "ROneCOne.File.WriteAllText path, \"hello files\"\nROneCOne.File.ReadAllText path",
        "hello files",
      ],
      [
        "Let a byte-order mark decide",
        "StreamReader detects encoding",
        "ROneCOne.File.WriteAllText path16, text, \"utf-16\"\nROneCOne.File.ReadAllText(path16) = text",
        true,
      ],
      [
        "Round-trip lines as a list",
        "File.ReadAllLines(path)",
        "ROneCOne.File.WriteAllLines path, Array(\"alpha\", \"beta\", \"gamma\")\nROneCOne.File.ReadAllLines(path).Count",
        3,
      ],
      [
        "Join path parts safely",
        "Path.Combine(\"C:\\data\", \"in\", \"file.txt\")",
        "ROneCOne.Path.Combine(\"C:\\data\", \"in\", \"file.txt\")",
        "C:\\data\\in\\file.txt",
      ],
      [
        "Dissect a path",
        "Path.GetFileNameWithoutExtension(path)",
        "ROneCOne.Path.GetFileNameWithoutExtension(\"C:\\data\\in\\file.txt\")",
        "file",
      ],
      [
        "Enumerate a tree with a pattern",
        "Directory.GetFiles(root, \"*.txt\", AllDirectories)",
        "ROneCOne.Directory.GetFiles(demoRoot, \"*.txt\", True).Count",
        4,
      ],
      [
        "Save a table as a CSV file",
        "File.WriteAllText(path, csv)",
        "ROneCOne.File.WriteAllText csvPath, orders.ToCsv\nROneCOne.File.Exists(csvPath)",
        true,
      ],
      [
        "Load the CSV file back",
        "CsvHelper-style typed parsing",
        "Set roundTripped = ROneCOne.Csv.DeserializeTable( _\n    ROneCOne.File.ReadAllText(csvPath), \"Orders\")",
        3,
      ],
      [
        "Columns come back typed",
        "deterministic type inference",
        "roundTripped.Rows.Item(0).Item(\"Total\")",
        12.5,
      ],
      [
        "Nulls survive the trip",
        "empty field vs quoted empty string",
        "IsNull(roundTripped.Rows.Item(0).Item(\"Note\"))",
        true,
      ],
    ],
  },
  {
    key: "process",
    title: "ROneCOne Processes",
    subtitle: "Await command lines with exit codes and captured output",
    macro: "RunROneCOneProcessDemo",
    feature: "Processes",
    output: "ROneCOne_Process_Demo.xlsx",
    benchmark: "Overlap three shell commands",
    benchmarkResult: "Collected results",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["Outside Excel", "WScript.Shell Exec", "Commands run beside Excel, never inside it", "ENFORCED", 1, 0],
      ["No pipe deadlock", "Scratch-file capture", "Output volume can never freeze Excel", "ENFORCED", 1, 0],
      ["Honest failures", "ExitCode + StandardError", "Failing commands report; only start failures raise", "ENFORCED", 1, 0],
      ["Cooperative await", "Status polling", "Excel stays responsive while commands run", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Run a command and await it",
        "Process.Start + WaitForExit",
        "Set hello = ROneCOne.Process.RunAsync(\"echo hello from ROneCOne\").Await\nhello.ExitCode = 0",
        true,
      ],
      [
        "Read captured output",
        "process.StandardOutput.ReadToEnd()",
        "InStr(1, hello.StandardOutput, \"hello from ROneCOne\") > 0",
        true,
      ],
      [
        "Exit codes pass through",
        "process.ExitCode",
        "ROneCOne.Process.RunAsync(\"exit 7\").Await.ExitCode",
        7,
      ],
      [
        "Standard error stays separate",
        "process.StandardError.ReadToEnd()",
        "InStr(1, warning.StandardError, \"be careful\") > 0",
        true,
      ],
      [
        "Pick the working directory",
        "startInfo.WorkingDirectory",
        "Set located = ROneCOne.Process.RunAsync(\"cd\", ThisWorkbook.Path).Await",
        true,
      ],
      [
        "Overlap commands with WhenAll",
        "Task.WhenAll(processes)",
        "ROneCOne.Task.WhenAll( _\n    ROneCOne.Process.RunAsync(\"echo first\"), _\n    ROneCOne.Process.RunAsync(\"echo second\")).Await.Count",
        2,
      ],
      [
        "Failing commands report",
        "nonzero ExitCode, no exception",
        "ROneCOne.Process.RunAsync(\"definitely_not_a_command_xyz\") _\n    .Await.ExitCode <> 0",
        true,
      ],
      [
        "Empty commands are rejected",
        "ArgumentException",
        "On Error Resume Next\nROneCOne.Process.RunAsync \"   \"\nIf Err.Number = ROneCOne.InvalidArgumentError Then ...",
        "empty command rejected",
      ],
    ],
  },
  {
    key: "text",
    title: "ROneCOne Text + Hashing",
    subtitle: "Match, format, build, hash, and encode text, all offline",
    macro: "RunROneCOneTextDemo",
    feature: "Text + hashing",
    output: "ROneCOne_Text_Demo.xlsx",
    benchmark: "Hash 1,000 strings with SHA-256",
    benchmarkResult: "Distinct digests",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["In-box regex", "VBScript.RegExp", "No reference, System.Text.RegularExpressions verbs", "ENFORCED", 1, 0],
      ["Native hashing", "Windows CNG (bcrypt.dll)", "No registration, FIPS and RFC vectors", "ENFORCED", 1, 0],
      ["Cross-platform digests", "Text hashed as UTF-8", "Matches Python, C#, and sha256sum", "ENFORCED", 1, 0],
      ["Offline", "No network", "Everything runs in-process", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Test whether text matches",
        "Regex.IsMatch(input, pattern)",
        "Set email = ROneCOne.Regex(\"(\\w+)@(\\w+)\\.(\\w+)\")\nemail.IsMatch(\"write ada@x.com today\")",
        true,
      ],
      [
        "Read a capture group",
        "match.Groups[1].Value",
        "email.Match(\"write ada@x.com today\").Groups.Item(1)",
        "ada",
      ],
      [
        "Count every match",
        "Regex.Matches(input).Count",
        "email.Matches(\"ada@x.com and bo@y.org\").Count",
        2,
      ],
      [
        "Replace with group references",
        "Regex.Replace(input, \"$2:$1\")",
        "email.Replace(\"ada@x.com\", \"$2:$1\")",
        "x:ada",
      ],
      [
        "Split on a pattern",
        "Regex.Split(input)",
        "ROneCOne.Regex(\"\\s*,\\s*\").Split(\"a, b ,c\").Count",
        3,
      ],
      [
        "Hash text as SHA-256",
        "SHA256.HashData(Encoding.UTF8.GetBytes(s))",
        "ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha256(\"abc\"))",
        "BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD",
      ],
      [
        "Sign with HMAC-SHA256",
        "new HMACSHA256(key).ComputeHash(msg)",
        "ROneCOne.Convert.ToHexString(ROneCOne.Hash.HmacSha256(\"Jefe\", _\n    \"what do ya want for nothing?\"))",
        "5BDCC146BF60754E6A042426089575C75A003F089D2739839DEC58B964EC3843",
      ],
      [
        "Encode bytes as base64",
        "Convert.ToBase64String(bytes)",
        "ROneCOne.Convert.ToBase64String(ROneCOne.Convert.FromHexString(\"4D616E\"))",
        "TWFu",
      ],
      [
        "Round-trip through hex",
        "Convert.ToHexString(bytes)",
        "ROneCOne.Convert.ToHexString(ROneCOne.Convert.FromHexString(\"4d616e\"))",
        "4D616E",
      ],
      [
        "Format values invariantly",
        "String.Format(\"{0} owes {1:N2}\", ...)",
        "ROneCOne.Strings.Format(\"{0} owes {1:N2}\", \"Ada\", 1234.5)",
        "Ada owes 1,234.50",
      ],
      [
        "Align, pad, and hex in one item",
        "composite format {0,6:X4}",
        "ROneCOne.Strings.Format(\"[{0,6:X4}]\", 255)",
        "[  00FF]",
      ],
      [
        "Build text in linear time",
        "new StringBuilder().Append(...)",
        "ROneCOne.StringBuilder().Append(\"a\").AppendFormat(\"{0:D3}\", 7).ToString",
        "a007",
      ],
      [
        "Mint ids and random bytes",
        "Guid.NewGuid / RandomNumberGenerator",
        "Len(ROneCOne.Guid.NewGuid) & \" chars, \" & _\n    (UBound(ROneCOne.RandomNumberGenerator.GetBytes(8)) + 1) & \" bytes\"",
        "36 chars, 8 bytes",
      ],
    ],
  },
  {
    key: "datetime",
    title: "ROneCOne Dates + Times",
    subtitle: "Parse ISO 8601 and epochs, convert zones, and add durations",
    macro: "RunROneCOneDateTimeDemo",
    feature: "Dates + times",
    output: "ROneCOne_DateTime_Demo.xlsx",
    benchmark: "Parse and format 1,000 timestamps",
    benchmarkResult: "Exact round trips",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["Windows owns the zones", "kernel32 conversion", "Daylight saving applied per instant, no tables in VBA", "ENFORCED", 1, 0],
      ["Instant model", "Epoch milliseconds + offset", "DateTimeOffset semantics at millisecond precision", "ENFORCED", 1, 0],
      ["Typed failures", "FormatError", "Impossible dates, hours, and offsets are refused", "ENFORCED", 1, 0],
      ["Offline", "No network", "Everything runs in-process", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Parse ISO 8601 with an offset",
        "DateTimeOffset.Parse(text)",
        "Set posted = ROneCOne.DateTime.Parse(\"2026-07-24T18:30:05.123+02:00\")\nposted.Hour",
        18,
      ],
      [
        "See the same instant in UTC",
        "value.ToUniversalTime()",
        "\"utc \" & posted.ToUniversalTime.ToIsoString",
        "utc 2026-07-24T16:30:05.123Z",
      ],
      [
        "Epoch numbers become readable",
        "DateTimeOffset.FromUnixTimeSeconds(n)",
        "\"epoch \" & ROneCOne.DateTime.FromUnixTimeSeconds(1784910605).ToIsoString",
        "epoch 2026-07-24T16:30:05Z",
      ],
      [
        "Readable times become epochs",
        "value.ToUnixTimeSeconds()",
        "CStr(posted.ToUnixTimeSeconds) & \" seconds\"",
        "1784910605 seconds",
      ],
      [
        "Different offsets, one instant",
        "CompareTo orders by instant",
        "posted.CompareTo(ROneCOne.DateTime.Parse( _\n    \"2026-07-24T16:30:05.123Z\")) = 0",
        true,
      ],
      [
        "Month ends clamp like .NET",
        "value.AddMonths(1)",
        "ROneCOne.DateTime.Parse(\"2026-01-31T12:00:00Z\") _\n    .AddMonths(1).ToIsoString = \"2026-02-28T12:00:00Z\"",
        true,
      ],
      [
        "Subtract instants into a duration",
        "later - earlier is a TimeSpan",
        "due.Subtract(startAt).TotalHours",
        2.5,
      ],
      [
        "Durations format themselves",
        "TimeSpan.ToString()",
        "\"lasts \" & ROneCOne.TimeSpan.FromMinutes(90).ToString",
        "lasts 01:30:00",
      ],
      [
        "Bad text is refused, typed",
        "FormatException",
        "On Error Resume Next\nROneCOne.DateTime.Parse \"2026-02-30\"\nIf Err.Number = ROneCOne.FormatError Then ...",
        "impossible date refused",
      ],
    ],
  },
  {
    key: "xml",
    title: "ROneCOne XML",
    subtitle: "Query XML with XPath and land it in typed DataTables",
    macro: "RunROneCOneXmlDemo",
    feature: "XML",
    output: "ROneCOne_Xml_Demo.xlsx",
    benchmark: "Extract 1,000 rows from XML",
    benchmarkResult: "Typed table rows",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["In-box parser", "MSXML2.DOMDocument.6.0", "Ships with Windows, no reference", "ENFORCED", 1, 0],
      ["Secure by default", "DTD prohibited, externals unresolved", "Hostile documents are refused, typed", "ENFORCED", 1, 0],
      ["Shared inference", "CSV column typing", "XML, CSV, and JSON agree on value shapes", "ENFORCED", 1, 0],
      ["Offline", "No network", "Everything runs in-process", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Parse a document",
        "XDocument.Parse(text)",
        "Set doc = ROneCOne.Xml.Parse(catalogXml)\ndoc.Name",
        "catalog",
      ],
      [
        "Read an attribute",
        "element.Attribute(\"id\").Value",
        "CLng(doc.Elements(\"book\").Item(0).GetAttribute(\"id\"))",
        1,
      ],
      [
        "XPath finds nodes anywhere",
        "doc.XPathSelectElements(\"//book\")",
        "doc.SelectNodes(\"//book\").Count",
        2,
      ],
      [
        "Predicates filter in the query",
        "//book[@id='2']/title",
        "doc.SelectSingleNode(\"//book[@id='2']/title\").Value",
        "Second & Third",
      ],
      [
        "Misses are Nothing, not errors",
        "XPathSelectElement returns null",
        "doc.SelectSingleNode(\"//missing\") Is Nothing",
        true,
      ],
      [
        "Namespaces map once",
        "XmlNamespaceManager prefixes",
        "Set feed = ROneCOne.Xml.Parse(feedXml, \"xmlns:p='urn:demo'\")\nfeed.SelectNodes(\"//p:item\").Count",
        2,
      ],
      [
        "XML lands as a typed table",
        "DataSet.ReadXml",
        "Set books = ROneCOne.Xml.DeserializeTable(catalogXml, \"Books\", \"//book\")\nbooks.Rows.Item(0).Item(\"price\")",
        10.5,
      ],
      [
        "Tables write themselves back",
        "DataTable.WriteXml",
        "InStr(1, books.ToXml(), \"<Books>\") > 0",
        true,
      ],
      [
        "Hostile documents are refused",
        "DTDs prohibited by default",
        "On Error Resume Next\nROneCOne.Xml.Parse \"<!DOCTYPE a []><a/>\"\nIf Err.Number = ROneCOne.XmlError Then ...",
        "doctype refused",
      ],
    ],
  },
  {
    key: "zip",
    title: "ROneCOne Zip",
    subtitle: "Open and make zip archives, pure VBA, offline",
    macro: "RunROneCOneZipDemo",
    feature: "Zip",
    output: "ROneCOne_Zip_Demo.xlsx",
    benchmark: "Inflate a 1,000-line archived file",
    benchmarkResult: "Summed value",
    architecture: [
      ["Single-file core", "ROneCOne.cls", "One import", "ENFORCED", 1, 0],
      ["No COM component", "Pure-VBA engine", "No reference, add-in, or Shell automation", "ENFORCED", 1, 0],
      ["Verified reads", "RFC 1951 inflate + CRC-32", "Every entry is checked against its own CRC", "ENFORCED", 1, 0],
      ["Safe extraction", "Directory-traversal guard", "Zip-slip names are refused before any write", "ENFORCED", 1, 0],
      ["Interoperable", "Compress-Archive both ways", "Reads .NET deflate, writes what Expand-Archive reads", "ENFORCED", 1, 0],
    ],
    examples: [
      [
        "Open an archive and count entries",
        "ZipFile.OpenRead(path).Entries",
        "Set archive = ROneCOne.ZipFile.OpenRead(madePath)\narchive.Entries.Count",
        2,
      ],
      [
        "Read an entry without extracting",
        "entry.Open() into a reader",
        "archive.GetEntry(\"readme.txt\").ReadAllText()",
        "read me first",
      ],
      [
        "An entry knows its size",
        "entry.Length",
        "archive.GetEntry(\"sub/data.csv\").Length",
        15,
      ],
      [
        "Stored entries are uncompressed",
        "entry.CompressedLength = entry.Length",
        "archive.GetEntry(\"sub/data.csv\").CompressedLength = _\n    archive.GetEntry(\"sub/data.csv\").Length",
        true,
      ],
      [
        "PowerShell zips, ROneCOne opens",
        "Compress-Archive then OpenRead",
        "ROneCOne.ZipFile.OpenRead(psMadePath).Entries.Count >= 2",
        true,
      ],
      [
        "And inflates its DEFLATE",
        "RFC 1951 inflate",
        "InStr(1, ROneCOne.ZipFile.OpenRead(psMadePath) _\n    .GetEntry(\"readme.txt\").ReadAllText(), \"read me first\") > 0",
        true,
      ],
      [
        "ROneCOne writes, PowerShell reads",
        "CreateFromDirectory then Expand-Archive",
        "ROneCOne.File.ReadAllText(ps_out & \"\\readme.txt\")",
        "read me first",
      ],
      [
        "Zip-slip names are refused",
        "directory-traversal guard",
        "On Error Resume Next\nROneCOne.ZipFile.ExtractToDirectory hostilePath, target\nIf Err.Number = ROneCOne.ZipError Then ...",
        "traversal refused",
      ],
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
    config.notice ||
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
  architecture.getRange("A6:F10").values = config.architecture || [
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
