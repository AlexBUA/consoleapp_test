using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// ---------------------------------------------------------------------------
// Defaults — overridden by CLI args (see usage below)
// ---------------------------------------------------------------------------
string excelPath    = @"Files\Data.xlsx";
string templatePath = @"Files\Template.docx";
string outputDir    = "Output";

// ---------------------------------------------------------------------------
// Parse optional command-line arguments
//   --excel    <path>   path to the .xlsx file
//   --template <path>   path to the .docx template
//   --output   <dir>    output directory
//
// Example:
//   dotnet run -- --excel Files\OtherData.xlsx --template Files\OtherTemplate.docx --output MyReports
// ---------------------------------------------------------------------------
for (int i = 0; i < args.Length - 1; i++)
{
    switch (args[i].ToLowerInvariant())
    {
        case "--excel"    or "-e": excelPath    = args[++i]; break;
        case "--template" or "-t": templatePath = args[++i]; break;
        case "--output"   or "-o": outputDir    = args[++i]; break;
    }
}

if (!File.Exists(excelPath))    { Console.WriteLine($"Error: '{excelPath}' not found.");    return; }
if (!File.Exists(templatePath)) { Console.WriteLine($"Error: '{templatePath}' not found."); return; }

// ---------------------------------------------------------------------------
// Clear (or create) the output directory so every run starts fresh
// ---------------------------------------------------------------------------
if (Directory.Exists(outputDir))
{
    int deleted = 0, locked = 0;
    foreach (var f in Directory.GetFiles(outputDir, "*.docx"))
    {
        try   { File.Delete(f); deleted++; }
        catch { Console.WriteLine($"  Warning: cannot delete '{Path.GetFileName(f)}' (file in use — close it first)"); locked++; }
    }
    if (deleted > 0 || locked > 0)
        Console.WriteLine($"Cleared {deleted} file(s) from '{outputDir}'" + (locked > 0 ? $" ({locked} skipped — in use)" : ""));
}
else
{
    Directory.CreateDirectory(outputDir);
}

// ---------------------------------------------------------------------------
// Read Excel data (skips fully-empty rows; row 1 = headers, ignored here)
// Placeholders in the template are $A$, $B$, $C$, $D$, $E$ — one per column.
// ---------------------------------------------------------------------------
var dataRows = ReadExcel(excelPath);

if (dataRows.Count == 0) { Console.WriteLine("No data rows found in Excel file."); return; }

Console.WriteLine($"Data rows: {dataRows.Count}");
Console.WriteLine($"Template : {templatePath}");
Console.WriteLine($"Output   : {Path.GetFullPath(outputDir)}\n");

for (int i = 0; i < dataRows.Count; i++)
{
    var row = dataRows[i];

    // Map  $A$ → row[0],  $B$ → row[1], …  $E$ → row[4]
    var replacements = new Dictionary<string, string>();
    for (int c = 0; c < row.Count; c++)
        replacements[$"${(char)('A' + c)}$"] = row[c];

    string first   = Sanitize(row.Count > 0 ? row[0] : "");
    string outPath = Path.Combine(outputDir, $"Report_{i + 1:D3}_{first}.docx");

    File.Copy(templatePath, outPath, overwrite: true);
    FillTemplate(outPath, replacements);

    Console.WriteLine($"  [{i + 1:D3}] {Path.GetFileName(outPath)}");
}

Console.WriteLine($"\nDone — {dataRows.Count} file(s) written to '{Path.GetFullPath(outputDir)}'");

// ---------------------------------------------------------------------------
// Excel reader — returns data rows only (row 1 header is skipped)
// ---------------------------------------------------------------------------
static List<List<string>> ReadExcel(string path)
{
    var rows = new List<List<string>>();

    using var wb = new XLWorkbook(path);
    var ws       = wb.Worksheets.First();
    var range    = ws.RangeUsed();
    if (range is null) return rows;

    var allRows = range.Rows().ToList();

    // Start from row index 1 to skip the header row
    for (int r = 1; r < allRows.Count; r++)
    {
        var cells   = Enumerable.Range(1, 5).Select(c => allRows[r].Cell(c).GetString()).ToList();
        bool hasAny = cells.Any(v => !string.IsNullOrWhiteSpace(v));
        if (hasAny) rows.Add(cells);
    }

    return rows;
}

// ---------------------------------------------------------------------------
// Word template filler — body, tables, headers and footers
// ---------------------------------------------------------------------------
static void FillTemplate(string docxPath, Dictionary<string, string> replacements)
{
    using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);

    foreach (var para in doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>().ToList())
        ReplaceParagraph(para, replacements);

    foreach (var hp in doc.MainDocumentPart.HeaderParts)
    {
        foreach (var para in hp.Header.Descendants<Paragraph>().ToList())
            ReplaceParagraph(para, replacements);
        hp.Header.Save();
    }

    foreach (var fp in doc.MainDocumentPart.FooterParts)
    {
        foreach (var para in fp.Footer.Descendants<Paragraph>().ToList())
            ReplaceParagraph(para, replacements);
        fp.Footer.Save();
    }

    doc.MainDocumentPart.Document.Save();
}

// ---------------------------------------------------------------------------
// Paragraph replacement — joins all runs first to handle Word's run-splitting,
// then replaces and collapses back to a single run with the first run's style.
// ---------------------------------------------------------------------------
static void ReplaceParagraph(Paragraph para, Dictionary<string, string> replacements)
{
    string combined = string.Concat(para.Descendants<Text>().Select(t => t.Text));

    if (!replacements.Keys.Any(combined.Contains)) return;

    foreach (var (placeholder, value) in replacements)
        combined = combined.Replace(placeholder, EscapeXml(value));

    var firstRun = para.Elements<Run>().FirstOrDefault();
    var runProps = firstRun?.RunProperties?.CloneNode(true) as RunProperties;

    foreach (var run in para.Elements<Run>().ToList())
        run.Remove();

    var newRun = new Run();
    if (runProps is not null) newRun.AppendChild(runProps);
    newRun.AppendChild(new Text(combined) { Space = SpaceProcessingModeValues.Preserve });
    para.AppendChild(newRun);
}

static string Sanitize(string name)
{
    if (string.IsNullOrWhiteSpace(name)) return "Unknown";
    foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
    return name.Length > 40 ? name[..40] : name;
}

static string EscapeXml(string s) =>
    s.Replace("&", "&amp;")
     .Replace("<", "&lt;")
     .Replace(">", "&gt;")
     .Replace("\"", "&quot;")
     .Replace("'", "&apos;");
