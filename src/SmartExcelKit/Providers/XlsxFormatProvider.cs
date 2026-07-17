using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;
using SmartExcelKit.Styles;
using System.IO.Compression;
using System.Xml.Linq;

namespace SmartExcelKit.Providers;

/// <summary>
/// High-performance format provider for standard Microsoft Excel OpenXML (.xlsx) files.
/// </summary>
public sealed class XlsxFormatProvider : IWorkbookFormatProvider
{
    private static readonly XNamespace nsSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace nsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <inheritdoc />
    public void Read(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

            // Read VBA Project if present
            var vbaEntry = archive.GetEntry("xl/vbaProject.bin");
            if (vbaEntry != null)
            {
                using var vStream = vbaEntry.Open();
                using var ms = new MemoryStream();
                vStream.CopyTo(ms);
                workbook.VbaProjectBytes = ms.ToArray();
            }
            else
            {
                workbook.VbaProjectBytes = null;
            }

            // 1. Read Shared Strings if present
            var sharedStrings = new List<string>();
            var sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsEntry != null)
            {
                using var sStream = sharedStringsEntry.Open();
                var sDoc = XDocument.Load(sStream);
                if (sDoc.Root != null)
                {
                    sharedStrings.AddRange(sDoc.Root.Elements(nsSpreadsheet + "si")
                        .Select(si => si.Element(nsSpreadsheet + "t")?.Value ?? string.Empty));
                }
            }

            // 2. Read Workbook Sheet Info and Relationships
            var sheetsMap = new List<(string Id, string Name)>();
            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new ParsingException("Invalid XLSX file: xl/workbook.xml not found.", "INVALID_XLSX");

            using (var wbStream = workbookEntry.Open())
            {
                var wbDoc = XDocument.Load(wbStream);
                var sheetsEl = wbDoc.Root?.Element(nsSpreadsheet + "sheets");
                if (sheetsEl != null)
                {
                    foreach (var sheetEl in sheetsEl.Elements(nsSpreadsheet + "sheet"))
                    {
                        string name = sheetEl.Attribute("name")?.Value ?? "Sheet";
                        string rId = sheetEl.Attribute(nsRelationship + "id")?.Value ?? string.Empty;
                        sheetsMap.Add((rId, name));
                    }
                }
            }

            // Read workbook relationships to resolve sheet paths
            var relsMap = new Dictionary<string, string>();
            var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (relsEntry != null)
            {
                using var relsStream = relsEntry.Open();
                var relsDoc = XDocument.Load(relsStream);
                if (relsDoc.Root != null)
                {
                    foreach (var relEl in relsDoc.Root.Elements())
                    {
                        string id = relEl.Attribute("Id")?.Value ?? string.Empty;
                        string target = relEl.Attribute("Target")?.Value ?? string.Empty;
                        relsMap[id] = target;
                    }
                }
            }

            // Clear existing sheets
            while (workbook.Worksheets.Count > 0)
            {
                workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
            }

            // 3. Read Worksheet Data
            foreach (var (rId, name) in sheetsMap)
            {
                if (!relsMap.TryGetValue(rId, out string? relativePath)) continue;

                // Adjust path (typically worksheets/sheet1.xml)
                string fullPath = relativePath.StartsWith("/") ? relativePath.Substring(1) : $"xl/{relativePath}";
                var wsEntry = archive.GetEntry(fullPath);
                if (wsEntry == null) continue;

                var worksheet = workbook.AddWorksheet(name);
                using var wsStream = wsEntry.Open();
                var wsDoc = XDocument.Load(wsStream);
                var sheetDataEl = wsDoc.Root?.Element(nsSpreadsheet + "sheetData");
                if (sheetDataEl == null) continue;

                foreach (var rowEl in sheetDataEl.Elements(nsSpreadsheet + "row"))
                {
                    foreach (var cellEl in rowEl.Elements(nsSpreadsheet + "c"))
                    {
                        string refAttr = cellEl.Attribute("r")?.Value ?? string.Empty;
                        if (string.IsNullOrEmpty(refAttr)) continue;

                        var address = CellAddress.Parse(refAttr);
                        var cell = worksheet.Cell(address);

                        string type = cellEl.Attribute("t")?.Value ?? "n"; // Default is numeric
                        string? formula = cellEl.Element(nsSpreadsheet + "f")?.Value;
                        string? valStr = cellEl.Element(nsSpreadsheet + "v")?.Value;

                        if (!string.IsNullOrEmpty(formula))
                        {
                            cell.Formula = formula;
                        }

                        if (type == "s" && valStr != null && int.TryParse(valStr, out int strIndex))
                        {
                            cell.Value = strIndex >= 0 && strIndex < sharedStrings.Count ? sharedStrings[strIndex] : string.Empty;
                        }
                        else if (type == "b" && valStr != null)
                        {
                            cell.Value = valStr == "1" || string.Equals(valStr, "true", StringComparison.OrdinalIgnoreCase);
                        }
                        else if (valStr != null)
                        {
                            if (double.TryParse(valStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d)) cell.Value = d;
                            else cell.Value = valStr;
                        }
                    }
                }

                // Read sheet protection
                var sheetProtectionEl = wsDoc.Root?.Element(nsSpreadsheet + "sheetProtection");
                if (sheetProtectionEl != null)
                {
                    worksheet.ProtectionPasswordHash = sheetProtectionEl.Attribute("password")?.Value;
                }

                // Read merged cells
                var mergeCellsEl = wsDoc.Root?.Element(nsSpreadsheet + "mergeCells");
                if (mergeCellsEl != null)
                {
                    foreach (var mergeCellEl in mergeCellsEl.Elements(nsSpreadsheet + "mergeCell"))
                    {
                        string refAttr = mergeCellEl.Attribute("ref")?.Value ?? string.Empty;
                        if (!string.IsNullOrEmpty(refAttr))
                        {
                            worksheet.MergeCells(ExcelRangeAddress.Parse(refAttr));
                        }
                    }
                }
            }
        }
        catch (Exception ex) when (ex is not SmartExcelException)
        {
            throw new ParsingException("Failed to read Excel OpenXML document.", "XLSX_READ_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public void Write(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

            // 1. Gather all unique string values for Shared Strings Table (SST)
            var sharedStrings = new List<string>();
            var stringMap = new Dictionary<string, int>();

            foreach (var sheet in workbook.Worksheets)
            {
                foreach (var cellKvp in sheet.RawCells)
                {
                    if (cellKvp.Value.Value is string s && string.IsNullOrEmpty(cellKvp.Value.Formula))
                    {
                        if (!stringMap.ContainsKey(s))
                        {
                            stringMap[s] = sharedStrings.Count;
                            sharedStrings.Add(s);
                        }
                    }
                }
            }

            // Write sharedStrings.xml
            if (sharedStrings.Count > 0)
            {
                var entry = archive.CreateEntry("xl/sharedStrings.xml");
                using var entryStream = entry.Open();
                var sDoc = new XDocument(
                    new XElement(nsSpreadsheet + "sst",
                        new XAttribute("count", sharedStrings.Count),
                        new XAttribute("uniqueCount", sharedStrings.Count),
                        sharedStrings.Select(s => new XElement(nsSpreadsheet + "si", new XElement(nsSpreadsheet + "t", s)))
                    )
                );
                sDoc.Save(entryStream);
            }

            // 2. Write Worksheets
            var wsRelationIds = new List<(string Id, string Name, string Path)>();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var sheet = workbook.Worksheets[i];
                string relId = $"rId{i + 1}";
                string wsPath = $"worksheets/sheet{i + 1}.xml";
                wsRelationIds.Add((relId, sheet.Name, wsPath));

                var entry = archive.CreateEntry($"xl/{wsPath}");
                using var entryStream = entry.Open();
                WriteWorksheetXml(entryStream, sheet, stringMap);
            }

            // 3. Write styles.xml (Minimal Styles implementation)
            var stylesEntry = archive.CreateEntry("xl/styles.xml");
            using (var stylesStream = stylesEntry.Open())
            {
                WriteStylesXml(stylesStream, workbook.StyleRegistry);
            }

            // Write VBA Project bin if present
            if (workbook.VbaProjectBytes != null)
            {
                var vbaEntry = archive.CreateEntry("xl/vbaProject.bin");
                using var vbaStream = vbaEntry.Open();
                vbaStream.Write(workbook.VbaProjectBytes, 0, workbook.VbaProjectBytes.Length);
            }

            // 4. Write xl/_rels/workbook.xml.rels
            var relsEntry = archive.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var relsStream = relsEntry.Open())
            {
                XNamespace nsRels = "http://schemas.openxmlformats.org/package/2006/relationships";
                var relsDoc = new XDocument(
                    new XElement(nsRels + "Relationships",
                        wsRelationIds.Select(r => new XElement(nsRels + "Relationship",
                            new XAttribute("Id", r.Id),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                            new XAttribute("Target", r.Path)
                        )),
                        new XElement(nsRels + "Relationship",
                            new XAttribute("Id", "rIdStyles"),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
                            new XAttribute("Target", "styles.xml")
                        ),
                        sharedStrings.Count > 0 ? new XElement(nsRels + "Relationship",
                            new XAttribute("Id", "rIdSharedStrings"),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"),
                            new XAttribute("Target", "sharedStrings.xml")
                        ) : null,
                        workbook.VbaProjectBytes != null ? new XElement(nsRels + "Relationship",
                            new XAttribute("Id", "rIdVba"),
                            new XAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject"),
                            new XAttribute("Target", "vbaProject.bin")
                        ) : null
                    )
                );
                relsDoc.Save(relsStream);
            }

            // 5. Write xl/workbook.xml
            var workbookEntry = archive.CreateEntry("xl/workbook.xml");
            using (var wbStream = workbookEntry.Open())
            {
                var wbDoc = new XDocument(
                    new XElement(nsSpreadsheet + "workbook",
                        new XAttribute(XNamespace.Xmlns + "r", nsRelationship.NamespaceName),
                        new XElement(nsSpreadsheet + "sheets",
                            wsRelationIds.Select(r => new XElement(nsSpreadsheet + "sheet",
                                new XAttribute("name", r.Name),
                                new XAttribute("sheetId", r.Id.Replace("rId", "")),
                                new XAttribute(nsRelationship + "id", r.Id)
                            ))
                        )
                    )
                );
                wbDoc.Save(wbStream);
            }

            // 6. Write _rels/.rels
            var rootRelsEntry = archive.CreateEntry("_rels/.rels");
            using (var rootRelsStream = rootRelsEntry.Open())
            {
                XNamespace nsRels = "http://schemas.openxmlformats.org/package/2006/relationships";
                var rootRelsDoc = new XDocument(
                    new XElement(nsRels + "Relationships",
                        new XElement(nsRels + "Relationship",
                            new XAttribute("Id", "rIdWb"),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                            new XAttribute("Target", "xl/workbook.xml")
                        )
                    )
                );
                rootRelsDoc.Save(rootRelsStream);
            }

            // 7. Write [Content_Types].xml
            var contentTypesEntry = archive.CreateEntry("[Content_Types].xml");
            using (var ctStream = contentTypesEntry.Open())
            {
                XNamespace nsTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
                var ctDoc = new XDocument(
                    new XElement(nsTypes + "Types",
                        new XElement(nsTypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                        new XElement(nsTypes + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                        workbook.VbaProjectBytes != null ? new XElement(nsTypes + "Default", new XAttribute("Extension", "bin"), new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")) : null,
                        new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/workbook.xml"), new XAttribute("ContentType", workbook.VbaProjectBytes != null ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                        wsRelationIds.Select(r => new XElement(nsTypes + "Override",
                            new XAttribute("PartName", $"/xl/{r.Path}"),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
                        )),
                        new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")),
                        sharedStrings.Count > 0 ? new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/sharedStrings.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")) : null,
                        workbook.VbaProjectBytes != null ? new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/vbaProject.bin"), new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")) : null
                    )
                );
                ctDoc.Save(ctStream);
            }
        }
        catch (Exception ex)
        {
            throw new ExportException("Failed to export Excel OpenXML archive.", "XLSX_WRITE_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public Task ReadAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Read(stream, workbook), cancellationToken);
    }

    /// <inheritdoc />
    public Task WriteAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(stream, workbook), cancellationToken);
    }

    private static void WriteWorksheetXml(Stream stream, ExcelWorksheet sheet, Dictionary<string, int> stringMap)
    {
        var wsEl = new XElement(nsSpreadsheet + "worksheet",
            new XAttribute("xmlns", nsSpreadsheet.NamespaceName),
            new XElement(nsSpreadsheet + "sheetData",
                // Sort cells by row, then column to produce sequential cell writing
                sheet.RawCells.GroupBy(c => c.Key.Row)
                    .OrderBy(g => g.Key)
                    .Select(g => new XElement(nsSpreadsheet + "row",
                        new XAttribute("r", g.Key),
                        g.OrderBy(c => c.Key.Column)
                         .Select(c =>
                         {
                             var cellVal = c.Value.Value;
                             var formula = c.Value.Formula;
                             var cEl = new XElement(nsSpreadsheet + "c", new XAttribute("r", c.Key.Address));

                             // Write custom style index if not default
                             if (c.Value.StyleId > 0)
                             {
                                 cEl.Add(new XAttribute("s", c.Value.StyleId));
                             }

                             if (!string.IsNullOrEmpty(formula))
                             {
                                 cEl.Add(new XElement(nsSpreadsheet + "f", formula));
                                 if (cellVal != null)
                                 {
                                     cEl.Add(new XElement(nsSpreadsheet + "v", FormatValue(cellVal)));
                                 }
                             }
                             else if (cellVal is string s)
                             {
                                 cEl.Add(new XAttribute("t", "s")); // sharedString type
                                 cEl.Add(new XElement(nsSpreadsheet + "v", stringMap[s]));
                             }
                             else if (cellVal is bool b)
                             {
                                 cEl.Add(new XAttribute("t", "b"));
                                 cEl.Add(new XElement(nsSpreadsheet + "v", b ? "1" : "0"));
                             }
                             else if (cellVal != null)
                             {
                                 cEl.Add(new XElement(nsSpreadsheet + "v", FormatValue(cellVal)));
                             }

                             return cEl;
                         })
                    ))
            )
        );

        if (sheet.IsProtected)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "sheetProtection",
                new XAttribute("password", sheet.ProtectionPasswordHash!),
                new XAttribute("sheet", "1"),
                new XAttribute("objects", "1"),
                new XAttribute("scenarios", "1")
            ));
        }

        if (sheet.MergedRanges.Count > 0)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "mergeCells",
                new XAttribute("count", sheet.MergedRanges.Count),
                sheet.MergedRanges.Select(r => new XElement(nsSpreadsheet + "mergeCell", new XAttribute("ref", r.Address)))
            ));
        }

        var doc = new XDocument(wsEl);
        doc.Save(stream);
    }

    private static void WriteStylesXml(Stream stream, StyleRegistry registry)
    {
        var doc = new XDocument(
            new XElement(nsSpreadsheet + "styleSheet",
                new XAttribute("xmlns", nsSpreadsheet.NamespaceName),
                new XElement(nsSpreadsheet + "fonts", new XAttribute("count", 1),
                    new XElement(nsSpreadsheet + "font",
                        new XElement(nsSpreadsheet + "sz", new XAttribute("val", 11)),
                        new XElement(nsSpreadsheet + "name", new XAttribute("val", "Calibri"))
                    )
                ),
                new XElement(nsSpreadsheet + "fills", new XAttribute("count", 2),
                    new XElement(nsSpreadsheet + "fill", new XElement(nsSpreadsheet + "patternFill", new XAttribute("patternType", "none"))),
                    new XElement(nsSpreadsheet + "fill", new XElement(nsSpreadsheet + "patternFill", new XAttribute("patternType", "gray125")))
                ),
                new XElement(nsSpreadsheet + "borders", new XAttribute("count", 1),
                    new XElement(nsSpreadsheet + "border",
                        new XElement(nsSpreadsheet + "left"),
                        new XElement(nsSpreadsheet + "right"),
                        new XElement(nsSpreadsheet + "top"),
                        new XElement(nsSpreadsheet + "bottom")
                    )
                ),
                new XElement(nsSpreadsheet + "cellStyleXfs", new XAttribute("count", 1),
                    new XElement(nsSpreadsheet + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0))
                ),
                // Cell formatting records map to the RegisteredStyles index in StyleRegistry
                new XElement(nsSpreadsheet + "cellXfs", new XAttribute("count", registry.RegisteredStyles.Count),
                    registry.RegisteredStyles.Select((style, index) => new XElement(nsSpreadsheet + "xf",
                        new XAttribute("numFmtId", 0),
                        new XAttribute("fontId", 0),
                        new XAttribute("fillId", 0),
                        new XAttribute("borderId", 0),
                        new XAttribute("xfId", 0),
                        new XAttribute("applyFont", style.Font.Equals(default) ? 0 : 1)
                    ))
                )
            )
        );
        doc.Save(stream);
    }

    private static string FormatValue(object val)
    {
        if (val is DateTime dt)
        {
            return dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        return Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
    }
}
