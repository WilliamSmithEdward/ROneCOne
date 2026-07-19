const fs = require("node:fs/promises");
const path = require("node:path");
const { FileBlob, SpreadsheetFile } = require("@oai/artifact-tool");

async function main() {
  const root = path.resolve(__dirname, "..");
  const workbookPath = path.join(root, "demo", "ROneCOne_Demo.xlsm");
  const outputDir = path.join(root, "demo", ".working");
  const input = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(input);

  const inspection = await workbook.inspect({
    kind: "sheet,formula,match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    maxChars: 4000,
    options: { maxResults: 100 },
  });
  await fs.writeFile(
    path.join(outputDir, "final-inspect.ndjson"),
    inspection.ndjson,
    "utf8",
  );

  for (const sheetName of ["Start Here", "Examples"]) {
    const preview = await workbook.render({
      sheetName,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    const safeName = sheetName.toLowerCase().replaceAll(" ", "-");
    await fs.writeFile(
      path.join(outputDir, `final-${safeName}.png`),
      new Uint8Array(await preview.arrayBuffer()),
    );
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
