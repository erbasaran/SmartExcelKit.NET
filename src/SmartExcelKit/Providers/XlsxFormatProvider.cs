using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;
using SmartExcelKit.PageSetup;
using SmartExcelKit.Styles;
using SmartExcelKit.Validation;

namespace SmartExcelKit.Providers;

/// <summary>
/// High-performance format provider for standard Microsoft Excel OpenXML (.xlsx) files.
/// Supports reading and writing cell values, formulas, styles, tables, conditional formatting, data validation, page setup, named ranges, drawings, charts, rich text, hyperlinks, comments, and protection.
/// </summary>
public sealed class XlsxFormatProvider : IWorkbookFormatProvider
{
    private static readonly XNamespace nsSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace nsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace nsChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace nsDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace nsMain = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
                        .Select(si =>
                        {
                            var tEl = si.Element(nsSpreadsheet + "t");
                            if (tEl != null) return tEl.Value;
                            return string.Concat(si.Elements(nsSpreadsheet + "r").Select(r => r.Element(nsSpreadsheet + "t")?.Value ?? string.Empty));
                        }));
                }
            }

            // 2. Read Workbook Sheet Info and Relationships
            var sheetsMap = new List<(string Id, string Name)>();
            var pendingNamedRanges = new List<(string Name, string RefersTo, string? LocalSheetId)>();
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

                // Read Workbook Protection
                var wbProtectionEl = wbDoc.Root?.Element(nsSpreadsheet + "workbookProtection");
                if (wbProtectionEl != null)
                {
                    workbook.ProtectionPasswordHash = wbProtectionEl.Attribute("workbookPassword")?.Value;
                }

                // Collect pending Named Ranges to resolve after worksheets are loaded
                var definedNamesEl = wbDoc.Root?.Element(nsSpreadsheet + "definedNames");
                if (definedNamesEl != null)
                {
                    foreach (var dnEl in definedNamesEl.Elements(nsSpreadsheet + "definedName"))
                    {
                        string name = dnEl.Attribute("name")?.Value ?? string.Empty;
                        string refersTo = dnEl.Value;
                        string? localSheetId = dnEl.Attribute("localSheetId")?.Value;

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(refersTo))
                        {
                            pendingNamedRanges.Add((name, refersTo, localSheetId));
                        }
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
            for (int sIdx = 0; sIdx < sheetsMap.Count; sIdx++)
            {
                var (rId, name) = sheetsMap[sIdx];
                if (!relsMap.TryGetValue(rId, out string? relativePath)) continue;

                string fullPath = relativePath.StartsWith("/") ? relativePath.Substring(1) : $"xl/{relativePath}";
                var wsEntry = archive.GetEntry(fullPath);
                if (wsEntry == null) continue;

                var worksheet = workbook.AddWorksheet(name);
                using var wsStream = wsEntry.Open();
                var wsDoc = XDocument.Load(wsStream);
                var sheetDataEl = wsDoc.Root?.Element(nsSpreadsheet + "sheetData");
                if (sheetDataEl != null)
                {
                    foreach (var rowEl in sheetDataEl.Elements(nsSpreadsheet + "row"))
                    {
                        foreach (var cellEl in rowEl.Elements(nsSpreadsheet + "c"))
                        {
                            string refAttr = cellEl.Attribute("r")?.Value ?? string.Empty;
                            if (string.IsNullOrEmpty(refAttr)) continue;

                            var address = CellAddress.Parse(refAttr);
                            var cell = worksheet.Cell(address);

                            string type = cellEl.Attribute("t")?.Value ?? "n";
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
                }

                // Read AutoFilter
                var autoFilterEl = wsDoc.Root?.Element(nsSpreadsheet + "autoFilter");
                if (autoFilterEl != null)
                {
                    string refAttr = autoFilterEl.Attribute("ref")?.Value ?? string.Empty;
                    if (!string.IsNullOrEmpty(refAttr))
                    {
                        worksheet.AutoFilter(refAttr);
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

                // Read Drawings & Charts
                ReadSheetDrawings(archive, sIdx + 1, worksheet);
            }

            // Assign pending named ranges to workbook or specific worksheet scopes
            foreach (var (pName, pRefersTo, pLocalSheetId) in pendingNamedRanges)
            {
                if (pLocalSheetId != null && int.TryParse(pLocalSheetId, out int sIdx) && sIdx >= 0 && sIdx < workbook.Worksheets.Count)
                {
                    workbook.Worksheets[sIdx].NamedRanges.Add(pName, pRefersTo);
                }
                else
                {
                    workbook.NamedRanges.Add(pName, pRefersTo);
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

            // 1. Gather all unique string values and RichText for Shared Strings Table (SST)
            var sharedStrings = new List<object>();
            var stringMap = new Dictionary<string, int>();

            foreach (var sheet in workbook.Worksheets)
            {
                foreach (var cellKvp in sheet.RawCells)
                {
                    var cellData = cellKvp.Value;

                    // Automatic Date Style assignment if no style set
                    if (cellData.Value is DateTime && cellData.StyleId == 0)
                    {
                        cellData.StyleId = workbook.StyleRegistry.Register(new ExcelStyle(numberFormat: new ExcelNumberFormat("yyyy-mm-dd")));
                    }

                    var richObj = cellData.RichText ?? (cellData.Value as RichText);
                    if (richObj != null && richObj.Runs.Count > 0)
                    {
                        string key = "RICHTEXT_" + richObj.ToString();
                        if (!stringMap.ContainsKey(key))
                        {
                            stringMap[key] = sharedStrings.Count;
                            sharedStrings.Add(richObj);
                        }
                    }
                    else if (cellData.Value is string s && string.IsNullOrEmpty(cellData.Formula))
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
                        sharedStrings.Select(item =>
                        {
                            if (item is RichText rt)
                            {
                                var siEl = new XElement(nsSpreadsheet + "si");
                                foreach (var run in rt.Runs)
                                {
                                    var rPr = new XElement(nsSpreadsheet + "rPr");
                                    if (run.Font.Bold) rPr.Add(new XElement(nsSpreadsheet + "b"));
                                    if (run.Font.Italic) rPr.Add(new XElement(nsSpreadsheet + "i"));
                                    if (run.Font.Underline) rPr.Add(new XElement(nsSpreadsheet + "u"));
                                    if (run.Font.Size > 0) rPr.Add(new XElement(nsSpreadsheet + "sz", new XAttribute("val", run.Font.Size)));
                                    if (!string.IsNullOrEmpty(run.Font.Color))
                                    {
                                        rPr.Add(new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(run.Font.Color!))));
                                    }
                                    if (!string.IsNullOrEmpty(run.Font.Name))
                                    {
                                        rPr.Add(new XElement(nsSpreadsheet + "rFont", new XAttribute("val", run.Font.Name)));
                                    }

                                    var rEl = new XElement(nsSpreadsheet + "r");
                                    if (rPr.HasElements) rEl.Add(rPr);
                                    rEl.Add(new XElement(nsSpreadsheet + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), run.Text));
                                    siEl.Add(rEl);
                                }
                                return siEl;
                            }
                            else
                            {
                                string s = item.ToString() ?? string.Empty;
                                return new XElement(nsSpreadsheet + "si", new XElement(nsSpreadsheet + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), s));
                            }
                        })
                    )
                );
                sDoc.Save(entryStream);
            }

            // 2. Write Worksheets & Drawings/Charts/Comments/Relationships
            var wsRelationIds = new List<(string Id, string Name, string Path)>();
            var contentTypeOverrides = new List<XElement>();
            int globalChartId = 1;
            int globalImageId = 1;

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var sheet = workbook.Worksheets[i];
                int sIdx = i + 1;
                string relId = $"rId{sIdx}";
                string wsPath = $"worksheets/sheet{sIdx}.xml";
                wsRelationIds.Add((relId, sheet.Name, wsPath));

                bool hasDrawing = sheet.Charts.Count > 0 || sheet.Images.Count > 0;
                var sheetRels = new List<(string Id, string Target, string? Mode)>();

                if (hasDrawing)
                {
                    WriteWorksheetDrawings(archive, sIdx, sheet, ref globalChartId, ref globalImageId, contentTypeOverrides, sheetRels);
                }

                // Check cell comments
                var commentCells = sheet.RawCells.Where(c => c.Value.CommentObject != null || !string.IsNullOrEmpty(c.Value.Comment)).ToList();
                bool hasComments = commentCells.Count > 0;
                if (hasComments)
                {
                    WriteSheetComments(archive, sIdx, commentCells, contentTypeOverrides, sheetRels);
                }

                var entry = archive.CreateEntry($"xl/{wsPath}");
                using (var entryStream = entry.Open())
                {
                    WriteWorksheetXml(entryStream, sheet, stringMap, hasDrawing, hasComments, sheetRels);
                }

                // Write xl/worksheets/_rels/sheet{sIdx}.xml.rels if sheetRels has items
                if (sheetRels.Count > 0)
                {
                    XNamespace nsRels = "http://schemas.openxmlformats.org/package/2006/relationships";
                    var sheetRelsDoc = new XDocument(
                        new XElement(nsRels + "Relationships",
                            sheetRels.Select(r =>
                            {
                                var relEl = new XElement(nsRels + "Relationship",
                                    new XAttribute("Id", r.Id),
                                    new XAttribute("Target", r.Target)
                                );

                                if (r.Id.StartsWith("rIdChart"))
                                {
                                    relEl.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"));
                                }
                                else if (r.Id.StartsWith("rIdImg"))
                                {
                                    relEl.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"));
                                }
                                else if (r.Id.StartsWith("rIdDrawing"))
                                {
                                    relEl.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"));
                                }
                                else if (r.Id.StartsWith("rIdHl"))
                                {
                                    relEl.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"));
                                    if (r.Mode != null) relEl.Add(new XAttribute("TargetMode", r.Mode));
                                }
                                else if (r.Id.StartsWith("rIdComment"))
                                {
                                    relEl.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"));
                                }
                                else if (r.Id.StartsWith("rIdVml"))
                                {
                                    relEl.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"));
                                }
                                return relEl;
                            })
                        )
                    );
                    var srEntry = archive.CreateEntry($"xl/worksheets/_rels/sheet{sIdx}.xml.rels");
                    using var srStream = srEntry.Open();
                    sheetRelsDoc.Save(srStream);
                }
            }

            // 3. Write styles.xml
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
                var sheetsEl = new XElement(nsSpreadsheet + "sheets",
                    wsRelationIds.Select(r => new XElement(nsSpreadsheet + "sheet",
                        new XAttribute("name", r.Name),
                        new XAttribute("sheetId", r.Id.Replace("rId", "")),
                        new XAttribute(nsRelationship + "id", r.Id)
                    ))
                );

                var wbEl = new XElement(nsSpreadsheet + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", nsRelationship.NamespaceName),
                    sheetsEl
                );

                if (workbook.IsProtected)
                {
                    wbEl.Add(new XElement(nsSpreadsheet + "workbookProtection",
                        new XAttribute("workbookPassword", workbook.ProtectionPasswordHash!),
                        new XAttribute("lockStructure", "1")
                    ));
                }

                var allDefinedNames = new List<XElement>();
                foreach (var nr in workbook.NamedRanges)
                {
                    allDefinedNames.Add(new XElement(nsSpreadsheet + "definedName",
                        new XAttribute("name", nr.Name),
                        nr.RefersTo
                    ));
                }

                for (int sIdx = 0; sIdx < workbook.Worksheets.Count; sIdx++)
                {
                    var sheet = workbook.Worksheets[sIdx];
                    foreach (var nr in sheet.NamedRanges)
                    {
                        allDefinedNames.Add(new XElement(nsSpreadsheet + "definedName",
                            new XAttribute("name", nr.Name),
                            new XAttribute("localSheetId", sIdx),
                            nr.RefersTo
                        ));
                    }
                }

                if (allDefinedNames.Count > 0)
                {
                    wbEl.Add(new XElement(nsSpreadsheet + "definedNames", allDefinedNames));
                }

                var wbDoc = new XDocument(wbEl);
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
                        new XElement(nsTypes + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")),
                        new XElement(nsTypes + "Default", new XAttribute("Extension", "jpeg"), new XAttribute("ContentType", "image/jpeg")),
                        new XElement(nsTypes + "Default", new XAttribute("Extension", "vml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing")),
                        workbook.VbaProjectBytes != null ? new XElement(nsTypes + "Default", new XAttribute("Extension", "bin"), new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")) : null,
                        new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/workbook.xml"), new XAttribute("ContentType", workbook.VbaProjectBytes != null ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                        wsRelationIds.Select(r => new XElement(nsTypes + "Override",
                            new XAttribute("PartName", $"/xl/{r.Path}"),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
                        )),
                        new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")),
                        sharedStrings.Count > 0 ? new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/sharedStrings.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")) : null,
                        workbook.VbaProjectBytes != null ? new XElement(nsTypes + "Override", new XAttribute("PartName", "/xl/vbaProject.bin"), new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")) : null,
                        contentTypeOverrides
                    )
                );
                ctDoc.Save(ctStream);
            }
        }
        catch (Exception ex)
        {
            throw new ExportException($"Failed to export Excel OpenXML archive: {ex.Message} -> {ex.InnerException?.Message}", "XLSX_WRITE_FAILED", ex);
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

    private static void WriteWorksheetXml(Stream stream, ExcelWorksheet sheet, Dictionary<string, int> stringMap, bool hasDrawing, bool hasComments, List<(string Id, string Target, string? Mode)> sheetRels)
    {
        var wsEl = new XElement(nsSpreadsheet + "worksheet",
            new XAttribute("xmlns", nsSpreadsheet.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", nsRelationship.NamespaceName)
        );

        // 1. Column Widths (<cols>)
        if (sheet.CustomColumnWidths.Count > 0)
        {
            var colsEl = new XElement(nsSpreadsheet + "cols",
                sheet.CustomColumnWidths.OrderBy(c => c.Key).Select(c => new XElement(nsSpreadsheet + "col",
                    new XAttribute("min", c.Key),
                    new XAttribute("max", c.Key),
                    new XAttribute("width", c.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                    new XAttribute("customWidth", "1")
                ))
            );
            wsEl.Add(colsEl);
        }

        // 2. Sheet Data (<sheetData>)
        var sheetDataEl = new XElement(nsSpreadsheet + "sheetData",
            sheet.RawCells.GroupBy(c => c.Key.Row)
                .OrderBy(g => g.Key)
                .Select(g =>
                {
                    var rowEl = new XElement(nsSpreadsheet + "row", new XAttribute("r", g.Key));
                    if (sheet.CustomRowHeights.TryGetValue(g.Key, out double rHt))
                    {
                        rowEl.Add(new XAttribute("ht", rHt.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                        rowEl.Add(new XAttribute("customHeight", "1"));
                    }

                    rowEl.Add(g.OrderBy(c => c.Key.Column)
                         .Select(c =>
                         {
                             var cellVal = c.Value.Value;
                             var formula = c.Value.Formula;
                             var cEl = new XElement(nsSpreadsheet + "c", new XAttribute("r", c.Key.Address));

                             if (c.Value.StyleId > 0)
                             {
                                 cEl.Add(new XAttribute("s", c.Value.StyleId));
                             }

                             var richObj = c.Value.RichText ?? (c.Value.Value as RichText);

                             if (!string.IsNullOrEmpty(formula))
                             {
                                 if (cellVal is string strResult)
                                 {
                                     cEl.Add(new XAttribute("t", "str"));
                                     cEl.Add(new XElement(nsSpreadsheet + "f", formula));
                                     cEl.Add(new XElement(nsSpreadsheet + "v", strResult));
                                 }
                                 else if (cellVal is bool boolResult)
                                 {
                                     cEl.Add(new XAttribute("t", "b"));
                                     cEl.Add(new XElement(nsSpreadsheet + "f", formula));
                                     cEl.Add(new XElement(nsSpreadsheet + "v", boolResult ? "1" : "0"));
                                 }
                                 else
                                 {
                                     cEl.Add(new XElement(nsSpreadsheet + "f", formula));
                                     if (cellVal != null)
                                     {
                                         cEl.Add(new XElement(nsSpreadsheet + "v", FormatValue(cellVal)));
                                     }
                                 }
                             }
                             else if (richObj != null && richObj.Runs.Count > 0)
                             {
                                 string key = "RICHTEXT_" + richObj.ToString();
                                 if (stringMap.TryGetValue(key, out int rIdx))
                                 {
                                     cEl.Add(new XAttribute("t", "s"));
                                     cEl.Add(new XElement(nsSpreadsheet + "v", rIdx));
                                 }
                             }
                             else if (cellVal is string s)
                             {
                                 cEl.Add(new XAttribute("t", "s"));
                                 cEl.Add(new XElement(nsSpreadsheet + "v", stringMap[s]));
                             }
                             else if (cellVal is bool b)
                             {
                                 cEl.Add(new XAttribute("t", "b"));
                                 cEl.Add(new XElement(nsSpreadsheet + "v", b ? "1" : "0"));
                             }
                             else if (cellVal is DateTime dt)
                             {
                                 cEl.Add(new XElement(nsSpreadsheet + "v", dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture)));
                             }
                             else if (cellVal != null)
                             {
                                 cEl.Add(new XElement(nsSpreadsheet + "v", FormatValue(cellVal)));
                             }

                             return cEl;
                         }));
                    return rowEl;
                })
        );
        wsEl.Add(sheetDataEl);

        // 3. Sheet Protection
        if (sheet.IsProtected)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "sheetProtection",
                new XAttribute("password", sheet.ProtectionPasswordHash!),
                new XAttribute("sheet", "1"),
                new XAttribute("objects", "1"),
                new XAttribute("scenarios", "1")
            ));
        }

        // 4. AutoFilter
        if (sheet.AutoFilterRange != null)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "autoFilter", new XAttribute("ref", sheet.AutoFilterRange.Value.Address)));
        }

        // 5. Merge Cells
        if (sheet.MergedRanges.Count > 0)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "mergeCells",
                new XAttribute("count", sheet.MergedRanges.Count),
                sheet.MergedRanges.Select(r => new XElement(nsSpreadsheet + "mergeCell", new XAttribute("ref", r.Address)))
            ));
        }

        // 6. Conditional Formatting
        if (sheet.ConditionalFormatting.Count > 0)
        {
            int priority = 1;
            foreach (var rule in sheet.ConditionalFormatting)
            {
                var cfRuleEl = new XElement(nsSpreadsheet + "cfRule",
                    new XAttribute("priority", priority++)
                );

                if (rule.StopIfTrue)
                {
                    cfRuleEl.Add(new XAttribute("stopIfTrue", "1"));
                }

                if (rule.RuleType == Formatting.ConditionalFormattingType.CellValue)
                {
                    cfRuleEl.Add(new XAttribute("type", "cellIs"));
                    string opStr = rule.Operator switch
                    {
                        Formatting.ConditionalFormattingOperator.Equal => "equal",
                        Formatting.ConditionalFormattingOperator.NotEqual => "notEqual",
                        Formatting.ConditionalFormattingOperator.GreaterThan => "greaterThan",
                        Formatting.ConditionalFormattingOperator.GreaterThanOrEqual => "greaterThanOrEqual",
                        Formatting.ConditionalFormattingOperator.LessThan => "lessThan",
                        Formatting.ConditionalFormattingOperator.LessThanOrEqual => "lessThanOrEqual",
                        Formatting.ConditionalFormattingOperator.Between => "between",
                        Formatting.ConditionalFormattingOperator.NotBetween => "notBetween",
                        Formatting.ConditionalFormattingOperator.ContainsText => "containsText",
                        Formatting.ConditionalFormattingOperator.StartsWith => "beginsWith",
                        Formatting.ConditionalFormattingOperator.EndsWith => "endsWith",
                        _ => "equal"
                    };
                    cfRuleEl.Add(new XAttribute("operator", opStr));

                    if (rule.Formula1 != null) cfRuleEl.Add(new XElement(nsSpreadsheet + "formula", rule.Formula1));
                    if (rule.Formula2 != null) cfRuleEl.Add(new XElement(nsSpreadsheet + "formula", rule.Formula2));
                }
                else if (rule.RuleType == Formatting.ConditionalFormattingType.Formula)
                {
                    cfRuleEl.Add(new XAttribute("type", "expression"));
                    if (rule.Formula1 != null) cfRuleEl.Add(new XElement(nsSpreadsheet + "formula", rule.Formula1));
                }
                else if (rule.RuleType == Formatting.ConditionalFormattingType.DataBar)
                {
                    cfRuleEl.Add(new XAttribute("type", "dataBar"));
                    string hex = FormatArgbHex(rule.Color1 ?? "00FF00");
                    var dataBarEl = new XElement(nsSpreadsheet + "dataBar",
                        new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "min")),
                        new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "max")),
                        new XElement(nsSpreadsheet + "color", new XAttribute("rgb", hex))
                    );
                    cfRuleEl.Add(dataBarEl);
                }
                else if (rule.RuleType == Formatting.ConditionalFormattingType.ColorScale)
                {
                    cfRuleEl.Add(new XAttribute("type", "colorScale"));
                    if (!string.IsNullOrEmpty(rule.Color3))
                    {
                        var colorScaleEl = new XElement(nsSpreadsheet + "colorScale",
                            new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "min")),
                            new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "percentile"), new XAttribute("val", "50")),
                            new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "max")),
                            new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(rule.Color1 ?? "FF0000"))),
                            new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(rule.Color2 ?? "FFFF00"))),
                            new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(rule.Color3)))
                        );
                        cfRuleEl.Add(colorScaleEl);
                    }
                    else
                    {
                        var colorScaleEl = new XElement(nsSpreadsheet + "colorScale",
                            new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "min")),
                            new XElement(nsSpreadsheet + "cfvo", new XAttribute("type", "max")),
                            new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(rule.Color1 ?? "FFFFFF"))),
                            new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(rule.Color2 ?? "0000FF")))
                        );
                        cfRuleEl.Add(colorScaleEl);
                    }
                }

                var cfEl = new XElement(nsSpreadsheet + "conditionalFormatting",
                    new XAttribute("sqref", rule.Range.Address),
                    cfRuleEl
                );
                wsEl.Add(cfEl);
            }
        }

        // 7. Data Validations
        if (sheet.DataValidations.Count > 0)
        {
            var dvsEl = new XElement(nsSpreadsheet + "dataValidations", new XAttribute("count", sheet.DataValidations.Count));
            foreach (var dv in sheet.DataValidations)
            {
                string vType = dv.ValidationType switch
                {
                    ValidationType.List => "list",
                    ValidationType.WholeNumber => "whole",
                    ValidationType.Decimal => "decimal",
                    ValidationType.Date => "date",
                    ValidationType.TextLength => "textLength",
                    ValidationType.Custom => "custom",
                    _ => "list"
                };

                string opStr = dv.Operator switch
                {
                    ValidationOperator.Between => "between",
                    ValidationOperator.NotBetween => "notBetween",
                    ValidationOperator.Equal => "equal",
                    ValidationOperator.NotEqual => "notEqual",
                    ValidationOperator.GreaterThan => "greaterThan",
                    ValidationOperator.GreaterThanOrEqual => "greaterThanOrEqual",
                    ValidationOperator.LessThan => "lessThan",
                    ValidationOperator.LessThanOrEqual => "lessThanOrEqual",
                    _ => "between"
                };

                var dvEl = new XElement(nsSpreadsheet + "dataValidation",
                    new XAttribute("type", vType),
                    new XAttribute("allowBlank", dv.AllowBlank ? "1" : "0"),
                    new XAttribute("showInputMessage", dv.ShowInputMessage ? "1" : "0"),
                    new XAttribute("showErrorMessage", dv.ShowErrorMessage ? "1" : "0"),
                    new XAttribute("sqref", dv.Range.Address)
                );

                if (dv.ValidationType != ValidationType.List && dv.ValidationType != ValidationType.Custom)
                {
                    dvEl.Add(new XAttribute("operator", opStr));
                }

                if (!string.IsNullOrEmpty(dv.ErrorTitle)) dvEl.Add(new XAttribute("errorTitle", dv.ErrorTitle));
                if (!string.IsNullOrEmpty(dv.ErrorMessage)) dvEl.Add(new XAttribute("error", dv.ErrorMessage));
                if (!string.IsNullOrEmpty(dv.PromptTitle)) dvEl.Add(new XAttribute("promptTitle", dv.PromptTitle));
                if (!string.IsNullOrEmpty(dv.Prompt)) dvEl.Add(new XAttribute("prompt", dv.Prompt));

                if (dv.Formula1 != null)
                {
                    string f1 = dv.Formula1;
                    if (dv.ValidationType == ValidationType.List && !f1.StartsWith("=") && !f1.StartsWith("\""))
                    {
                        f1 = $"\"{f1}\"";
                    }
                    dvEl.Add(new XElement(nsSpreadsheet + "formula1", f1));
                }
                if (dv.Formula2 != null) dvEl.Add(new XElement(nsSpreadsheet + "formula2", dv.Formula2));

                dvsEl.Add(dvEl);
            }
            wsEl.Add(dvsEl);
        }

        // 8. Hyperlinks
        var hyperlinks = sheet.RawCells.Where(c => c.Value.HyperlinkObject != null || !string.IsNullOrEmpty(c.Value.Hyperlink)).ToList();
        if (hyperlinks.Count > 0)
        {
            var hlsEl = new XElement(nsSpreadsheet + "hyperlinks");
            int hlRelIdx = 1;
            foreach (var kvp in hyperlinks)
            {
                var hl = kvp.Value.HyperlinkObject ?? new ExcelHyperlink(kvp.Value.Hyperlink!);
                var hlEl = new XElement(nsSpreadsheet + "hyperlink", new XAttribute("ref", kvp.Key.Address));
                bool isExternal = hl.HyperlinkType != HyperlinkType.InternalReference;
                if (isExternal)
                {
                    string rId = $"rIdHl{hlRelIdx++}";
                    hlEl.Add(new XAttribute(nsRelationship + "id", rId));
                    sheetRels.Add((rId, hl.Target, "External"));
                }
                else
                {
                    hlEl.Add(new XAttribute("location", hl.Target));
                }
                if (!string.IsNullOrEmpty(hl.Tooltip))
                {
                    hlEl.Add(new XAttribute("tooltip", hl.Tooltip));
                }
                hlsEl.Add(hlEl);
            }
            wsEl.Add(hlsEl);
        }

        // 9. Page Setup & Printing
        var ps = sheet.PageSetup;
        if (ps != null)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "printOptions",
                new XAttribute("gridLines", ps.PrintGridlines ? "1" : "0"),
                new XAttribute("headings", ps.PrintHeadings ? "1" : "0")
            ));

            wsEl.Add(new XElement(nsSpreadsheet + "pageMargins",
                new XAttribute("left", "0.7"),
                new XAttribute("right", "0.7"),
                new XAttribute("top", "0.75"),
                new XAttribute("bottom", "0.75"),
                new XAttribute("header", "0.3"),
                new XAttribute("footer", "0.3")
            ));

            string orient = ps.Orientation == PageOrientation.Landscape ? "landscape" : "portrait";
            int paperSize = (int)ps.PaperSize;
            if (paperSize <= 0) paperSize = 9;

            wsEl.Add(new XElement(nsSpreadsheet + "pageSetup",
                new XAttribute("paperSize", paperSize),
                new XAttribute("orientation", orient)
            ));

            if (!string.IsNullOrEmpty(ps.HeaderCenter) || !string.IsNullOrEmpty(ps.HeaderLeft) || !string.IsNullOrEmpty(ps.HeaderRight) ||
                !string.IsNullOrEmpty(ps.FooterCenter) || !string.IsNullOrEmpty(ps.FooterLeft) || !string.IsNullOrEmpty(ps.FooterRight))
            {
                string headerStr = $"&L{ps.HeaderLeft}&C{ps.HeaderCenter}&R{ps.HeaderRight}";
                string footerStr = $"&L{ps.FooterLeft}&C{ps.FooterCenter}&R{ps.FooterRight}";
                wsEl.Add(new XElement(nsSpreadsheet + "headerFooter",
                    new XElement(nsSpreadsheet + "oddHeader", headerStr),
                    new XElement(nsSpreadsheet + "oddFooter", footerStr)
                ));
            }
        }

        // 10. Drawings & Legacy VML Drawings (Comments)
        if (hasDrawing)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "drawing", new XAttribute(nsRelationship + "id", "rIdDrawing1")));
        }

        if (hasComments)
        {
            wsEl.Add(new XElement(nsSpreadsheet + "legacyDrawing", new XAttribute(nsRelationship + "id", "rIdVml1")));
        }

        var doc = new XDocument(wsEl);
        doc.Save(stream);
    }

    private static void WriteSheetComments(ZipArchive archive, int sIdx, List<KeyValuePair<CellAddress, CellData>> commentCells, List<XElement> contentTypeOverrides, List<(string Id, string Target, string? Mode)> sheetRels)
    {
        string commentPath = $"xl/comments{sIdx}.xml";
        string vmlPath = $"xl/drawings/vmlDrawing{sIdx}.vml";

        contentTypeOverrides.Add(new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Override",
            new XAttribute("PartName", $"/{commentPath}"),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
        ));

        sheetRels.Add(("rIdComment1", $"../comments{sIdx}.xml", null));
        sheetRels.Add(("rIdVml1", $"../drawings/vmlDrawing{sIdx}.vml", null));

        // 1. Write xl/comments{sIdx}.xml
        var authors = commentCells.Select(c => c.Value.CommentObject?.Author ?? "Author").Distinct().ToList();
        var commentsDoc = new XDocument(
            new XElement(nsSpreadsheet + "comments",
                new XAttribute("xmlns", nsSpreadsheet.NamespaceName),
                new XElement(nsSpreadsheet + "authors",
                    authors.Select(a => new XElement(nsSpreadsheet + "author", a))
                ),
                new XElement(nsSpreadsheet + "commentList",
                    commentCells.Select(c =>
                    {
                        string author = c.Value.CommentObject?.Author ?? "Author";
                        int aIdx = authors.IndexOf(author);
                        string text = c.Value.CommentObject?.Text ?? c.Value.Comment ?? string.Empty;

                        return new XElement(nsSpreadsheet + "comment",
                            new XAttribute("ref", c.Key.Address),
                            new XAttribute("authorId", Math.Max(0, aIdx)),
                            new XElement(nsSpreadsheet + "text",
                                new XElement(nsSpreadsheet + "r",
                                    new XElement(nsSpreadsheet + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)
                                )
                            )
                        );
                    })
                )
            )
        );
        var cEntry = archive.CreateEntry(commentPath);
        using (var cStream = cEntry.Open()) commentsDoc.Save(cStream);

        // 2. Write xl/drawings/vmlDrawing{sIdx}.vml
        var vmlSb = new StringBuilder();
        vmlSb.AppendLine("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
        vmlSb.AppendLine(" <o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
        vmlSb.AppendLine(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,z\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");

        int shapeId = 1025;
        foreach (var kvp in commentCells)
        {
            int r = kvp.Key.Row - 1;
            int c = kvp.Key.Column - 1;
            vmlSb.AppendLine($" <v:shape id=\"_x0000_s{shapeId++}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
            vmlSb.AppendLine("  <v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/>");
            vmlSb.AppendLine("  <x:ClientData ObjectType=\"Note\">");
            vmlSb.AppendLine("   <x:MoveWithCells/><x:SizeWithCells/>");
            vmlSb.AppendLine($"   <x:Anchor>{c}, 15, {r}, 2, {c + 2}, 31, {r + 4}, 1</x:Anchor>");
            vmlSb.AppendLine("   <x:AutoFill>False</x:AutoFill>");
            vmlSb.AppendLine($"   <x:Row>{r}</x:Row>");
            vmlSb.AppendLine($"   <x:Column>{c}</x:Column>");
            vmlSb.AppendLine("  </x:ClientData>");
            vmlSb.AppendLine(" </v:shape>");
        }
        vmlSb.AppendLine("</xml>");

        var vmlEntry = archive.CreateEntry(vmlPath);
        using var vmlWriter = new StreamWriter(vmlEntry.Open(), Encoding.UTF8);
        vmlWriter.Write(vmlSb.ToString());
    }

    private static void WriteWorksheetDrawings(ZipArchive archive, int sIdx, ExcelWorksheet sheet, ref int globalChartId, ref int globalImageId, List<XElement> contentTypeOverrides, List<(string Id, string Target, string? Mode)> sheetRels)
    {
        contentTypeOverrides.Add(new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Override",
            new XAttribute("PartName", $"/xl/drawings/drawing{sIdx}.xml"),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")
        ));

        sheetRels.Add(("rIdDrawing1", $"../drawings/drawing{sIdx}.xml", null));

        var drawingRels = new List<XElement>();
        var twoCellAnchors = new List<XElement>();

        // Write Charts
        foreach (var chart in sheet.Charts)
        {
            int cId = globalChartId++;
            string chartRelId = $"rIdChart{cId}";

            drawingRels.Add(new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships") + "Relationship",
                new XAttribute("Id", chartRelId),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                new XAttribute("Target", $"../charts/chart{cId}.xml")
            ));

            contentTypeOverrides.Add(new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Override",
                new XAttribute("PartName", $"/xl/charts/chart{cId}.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            ));

            WriteChartXml(archive, $"xl/charts/chart{cId}.xml", chart);

            int fromCol = Math.Max(0, chart.LeftColumn - 1);
            int fromRow = Math.Max(0, chart.TopRow - 1);
            int toCol = fromCol + Math.Max(2, chart.Width / 64);
            int toRow = fromRow + Math.Max(4, chart.Height / 20);

            twoCellAnchors.Add(new XElement(nsDrawing + "twoCellAnchor",
                new XAttribute("editAs", "twoCell"),
                new XElement(nsDrawing + "from",
                    new XElement(nsDrawing + "col", fromCol),
                    new XElement(nsDrawing + "colOff", 0),
                    new XElement(nsDrawing + "row", fromRow),
                    new XElement(nsDrawing + "rowOff", 0)
                ),
                new XElement(nsDrawing + "to",
                    new XElement(nsDrawing + "col", toCol),
                    new XElement(nsDrawing + "colOff", 0),
                    new XElement(nsDrawing + "row", toRow),
                    new XElement(nsDrawing + "rowOff", 0)
                ),
                new XElement(nsDrawing + "graphicFrame",
                    new XAttribute("macro", ""),
                    new XElement(nsDrawing + "nvGraphicFramePr",
                        new XElement(nsDrawing + "cNvPr", new XAttribute("id", cId + 1), new XAttribute("name", $"Chart {cId}")),
                        new XElement(nsDrawing + "cNvGraphicFramePr")
                    ),
                    new XElement(nsDrawing + "xfrm",
                        new XElement(nsMain + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(nsMain + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0))
                    ),
                    new XElement(nsMain + "graphic",
                        new XElement(nsMain + "graphicData", new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                            new XElement(nsChart + "chart",
                                new XAttribute(XNamespace.Xmlns + "c", nsChart.NamespaceName),
                                new XAttribute(nsRelationship + "id", chartRelId)
                            )
                        )
                    )
                ),
                new XElement(nsDrawing + "clientData")
            ));
        }

        // Write Images
        foreach (var image in sheet.Images)
        {
            int imgId = globalImageId++;
            string imgRelId = $"rIdImg{imgId}";
            string ext = image.Format.ToString().ToLowerInvariant();
            string mediaPath = $"media/image{imgId}.{ext}";

            drawingRels.Add(new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships") + "Relationship",
                new XAttribute("Id", imgRelId),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                new XAttribute("Target", $"../{mediaPath}")
            ));

            var imgEntry = archive.CreateEntry($"xl/{mediaPath}");
            using (var imgStream = imgEntry.Open())
            {
                imgStream.Write(image.ImageBytes, 0, image.ImageBytes.Length);
            }

            int fromCol = Math.Max(0, image.LeftColumn - 1);
            int fromRow = Math.Max(0, image.TopRow - 1);
            int toCol = fromCol + Math.Max(2, image.Width / 64);
            int toRow = fromRow + Math.Max(4, image.Height / 20);

            twoCellAnchors.Add(new XElement(nsDrawing + "twoCellAnchor",
                new XAttribute("editAs", "twoCell"),
                new XElement(nsDrawing + "from",
                    new XElement(nsDrawing + "col", fromCol),
                    new XElement(nsDrawing + "colOff", 0),
                    new XElement(nsDrawing + "row", fromRow),
                    new XElement(nsDrawing + "rowOff", 0)
                ),
                new XElement(nsDrawing + "to",
                    new XElement(nsDrawing + "col", toCol),
                    new XElement(nsDrawing + "colOff", 0),
                    new XElement(nsDrawing + "row", toRow),
                    new XElement(nsDrawing + "rowOff", 0)
                ),
                new XElement(nsDrawing + "pic",
                    new XElement(nsDrawing + "nvPicPr",
                        new XElement(nsDrawing + "cNvPr", new XAttribute("id", imgId + 100), new XAttribute("name", image.Name)),
                        new XElement(nsDrawing + "cNvPicPr")
                    ),
                    new XElement(nsDrawing + "blipFill",
                        new XElement(nsMain + "blip", new XAttribute(nsRelationship + "embed", imgRelId)),
                        new XElement(nsMain + "stretch", new XElement(nsMain + "fillRect"))
                    ),
                    new XElement(nsDrawing + "spPr",
                        new XElement(nsMain + "xfrm",
                            new XElement(nsMain + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                            new XElement(nsMain + "ext", new XAttribute("cx", image.Width * 9525), new XAttribute("cy", image.Height * 9525))
                        ),
                        new XElement(nsMain + "prstGeom", new XAttribute("prst", "rect"), new XElement(nsMain + "avLst"))
                    )
                ),
                new XElement(nsDrawing + "clientData")
            ));
        }

        // xl/drawings/drawing{sIdx}.xml
        var drDoc = new XDocument(
            new XElement(nsDrawing + "wsDr",
                new XAttribute(XNamespace.Xmlns + "xdr", nsDrawing.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "a", nsMain.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", nsRelationship.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "c", nsChart.NamespaceName),
                twoCellAnchors
            )
        );
        var drEntry = archive.CreateEntry($"xl/drawings/drawing{sIdx}.xml");
        using (var drStream = drEntry.Open()) drDoc.Save(drStream);

        // xl/drawings/_rels/drawing{sIdx}.xml.rels
        if (drawingRels.Count > 0)
        {
            var drRelsDoc = new XDocument(
                new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships") + "Relationships", drawingRels)
            );
            var drRelsEntry = archive.CreateEntry($"xl/drawings/_rels/drawing{sIdx}.xml.rels");
            using (var drRelsStream = drRelsEntry.Open()) drRelsDoc.Save(drRelsStream);
        }
    }

    private static string FormatChartRange(string range)
    {
        if (string.IsNullOrWhiteSpace(range)) return range;
        string sheetPart = string.Empty;
        string rangePart = range;
        int exIndex = range.IndexOf('!');
        if (exIndex >= 0)
        {
            sheetPart = range.Substring(0, exIndex).Trim('\'');
            rangePart = range.Substring(exIndex + 1);
        }

        try
        {
            var addr = ExcelRangeAddress.Parse(rangePart);
            bool isSingleCell = addr.StartColumn == addr.EndColumn && addr.StartRow == addr.EndRow;
            string formattedRange = isSingleCell
                ? $"${CellAddress.GetColumnName(addr.StartColumn)}${addr.StartRow}"
                : $"${CellAddress.GetColumnName(addr.StartColumn)}${addr.StartRow}:${CellAddress.GetColumnName(addr.EndColumn)}${addr.EndRow}";

            return string.IsNullOrEmpty(sheetPart) ? formattedRange : $"'{sheetPart}'!{formattedRange}";
        }
        catch
        {
            return range;
        }
    }

    private static void WriteChartXml(ZipArchive archive, string chartPath, Drawings.ExcelChart chart)
    {
        string chartTypeTag = chart.ChartType switch
        {
            Drawings.ChartType.Column or Drawings.ChartType.Bar => "barChart",
            Drawings.ChartType.Line => "lineChart",
            Drawings.ChartType.Pie => "pieChart",
            Drawings.ChartType.Doughnut => "doughnutChart",
            Drawings.ChartType.Area => "areaChart",
            Drawings.ChartType.Scatter => "scatterChart",
            _ => "barChart"
        };

        var seriesEls = new List<XElement>();
        for (int i = 0; i < chart.Series.Count; i++)
        {
            var s = chart.Series[i];
            var sEl = new XElement(nsChart + "ser",
                new XElement(nsChart + "idx", new XAttribute("val", i)),
                new XElement(nsChart + "order", new XAttribute("val", i)),
                new XElement(nsChart + "tx", new XElement(nsChart + "v", s.Name))
            );

            if (!string.IsNullOrEmpty(chart.CategoryRange))
            {
                sEl.Add(new XElement(nsChart + "cat",
                    new XElement(nsChart + "strRef",
                        new XElement(nsChart + "f", FormatChartRange(chart.CategoryRange!))
                    )
                ));
            }

            sEl.Add(new XElement(nsChart + "val",
                new XElement(nsChart + "numRef",
                    new XElement(nsChart + "f", FormatChartRange(s.ValuesRange))
                )
            ));

            seriesEls.Add(sEl);
        }

        var chartTypeEl = new XElement(nsChart + chartTypeTag);
        chartTypeEl.Add(new XElement(nsChart + "varyColors", new XAttribute("val", "0")));

        if (chart.ChartType == Drawings.ChartType.Column)
        {
            chartTypeEl.Add(new XElement(nsChart + "barDir", new XAttribute("val", "col")));
            chartTypeEl.Add(new XElement(nsChart + "grouping", new XAttribute("val", "standard")));
        }
        else if (chart.ChartType == Drawings.ChartType.Bar)
        {
            chartTypeEl.Add(new XElement(nsChart + "barDir", new XAttribute("val", "bar")));
            chartTypeEl.Add(new XElement(nsChart + "grouping", new XAttribute("val", "standard")));
        }
        else if (chart.ChartType == Drawings.ChartType.Line)
        {
            chartTypeEl.Add(new XElement(nsChart + "grouping", new XAttribute("val", "standard")));
        }

        foreach (var sEl in seriesEls) chartTypeEl.Add(sEl);

        chartTypeEl.Add(new XElement(nsChart + "axId", new XAttribute("val", "100001")));
        chartTypeEl.Add(new XElement(nsChart + "axId", new XAttribute("val", "100002")));

        string legendPos = chart.LegendPosition switch
        {
            Drawings.LegendPosition.Top => "t",
            Drawings.LegendPosition.Bottom => "b",
            Drawings.LegendPosition.Left => "l",
            Drawings.LegendPosition.Right => "r",
            Drawings.LegendPosition.TopRight => "tr",
            _ => "r"
        };

        var plotAreaEl = new XElement(nsChart + "plotArea",
            new XElement(nsChart + "layout"),
            chartTypeEl
        );

        if (chart.ChartType != Drawings.ChartType.Pie && chart.ChartType != Drawings.ChartType.Doughnut)
        {
            plotAreaEl.Add(new XElement(nsChart + "catAx",
                new XElement(nsChart + "axId", new XAttribute("val", "100001")),
                new XElement(nsChart + "scaling", new XElement(nsChart + "orientation", new XAttribute("val", "minMax"))),
                new XElement(nsChart + "delete", new XAttribute("val", "0")),
                new XElement(nsChart + "axPos", new XAttribute("val", "b")),
                new XElement(nsChart + "numFmt", new XAttribute("formatCode", "General"), new XAttribute("sourceLinked", "1")),
                new XElement(nsChart + "majorTickMark", new XAttribute("val", "out")),
                new XElement(nsChart + "minorTickMark", new XAttribute("val", "none")),
                new XElement(nsChart + "tickLblPos", new XAttribute("val", "nextTo")),
                new XElement(nsChart + "crossAx", new XAttribute("val", "100002")),
                new XElement(nsChart + "crosses", new XAttribute("val", "autoZero")),
                new XElement(nsChart + "auto", new XAttribute("val", "1")),
                new XElement(nsChart + "lblAlgn", new XAttribute("val", "ctr")),
                new XElement(nsChart + "lblOffset", new XAttribute("val", "100"))
            ));
            plotAreaEl.Add(new XElement(nsChart + "valAx",
                new XElement(nsChart + "axId", new XAttribute("val", "100002")),
                new XElement(nsChart + "scaling", new XElement(nsChart + "orientation", new XAttribute("val", "minMax"))),
                new XElement(nsChart + "delete", new XAttribute("val", "0")),
                new XElement(nsChart + "axPos", new XAttribute("val", "l")),
                new XElement(nsChart + "majorGridlines"),
                new XElement(nsChart + "numFmt", new XAttribute("formatCode", "General"), new XAttribute("sourceLinked", "1")),
                new XElement(nsChart + "majorTickMark", new XAttribute("val", "out")),
                new XElement(nsChart + "minorTickMark", new XAttribute("val", "none")),
                new XElement(nsChart + "tickLblPos", new XAttribute("val", "nextTo")),
                new XElement(nsChart + "crossAx", new XAttribute("val", "100001")),
                new XElement(nsChart + "crosses", new XAttribute("val", "autoZero"))
            ));
        }

        var chartSpaceDoc = new XDocument(
            new XElement(nsChart + "chartSpace",
                new XAttribute(XNamespace.Xmlns + "c", nsChart.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "a", nsMain.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", nsRelationship.NamespaceName),
                new XElement(nsChart + "date1904", new XAttribute("val", "0")),
                new XElement(nsChart + "lang", new XAttribute("val", "en-US")),
                new XElement(nsChart + "roundedCorners", new XAttribute("val", "0")),
                new XElement(nsChart + "chart",
                    new XElement(nsChart + "title",
                        new XElement(nsChart + "tx",
                            new XElement(nsChart + "rich",
                                new XElement(nsMain + "bodyPr"),
                                new XElement(nsMain + "lstStyle"),
                                new XElement(nsMain + "p",
                                    new XElement(nsMain + "pPr", new XElement(nsMain + "defRPr")),
                                    new XElement(nsMain + "r",
                                        new XElement(nsMain + "rPr", new XAttribute("lang", "en-US")),
                                        new XElement(nsMain + "t", chart.Title)
                                    )
                                )
                            )
                        ),
                        new XElement(nsChart + "layout"),
                        new XElement(nsChart + "overlay", new XAttribute("val", "0"))
                    ),
                    new XElement(nsChart + "autoTitleDeleted", new XAttribute("val", "0")),
                    plotAreaEl,
                    chart.LegendPosition != Drawings.LegendPosition.None ? new XElement(nsChart + "legend",
                        new XElement(nsChart + "legendPos", new XAttribute("val", legendPos)),
                        new XElement(nsChart + "layout"),
                        new XElement(nsChart + "overlay", new XAttribute("val", "0"))
                    ) : null,
                    new XElement(nsChart + "plotVisOnly", new XAttribute("val", "1")),
                    new XElement(nsChart + "dispBlanksAs", new XAttribute("val", "gap")),
                    new XElement(nsChart + "showDLblsOverMax", new XAttribute("val", "0"))
                )
            )
        );

        var cEntry = archive.CreateEntry(chartPath);
        using var cStream = cEntry.Open();
        chartSpaceDoc.Save(cStream);
    }

    private static void ReadSheetDrawings(ZipArchive archive, int sheetIndex, ExcelWorksheet worksheet)
    {
        var srEntry = archive.GetEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels");
        if (srEntry == null) return;

        using var srStream = srEntry.Open();
        var srDoc = XDocument.Load(srStream);
        var drawingRel = srDoc.Root?.Elements().FirstOrDefault(e => e.Attribute("Type")?.Value.EndsWith("drawing") == true);
        if (drawingRel == null) return;

        string drTarget = drawingRel.Attribute("Target")?.Value ?? string.Empty;
        string drPath = drTarget.StartsWith("/") ? drTarget.Substring(1) : (drTarget.StartsWith("../") ? $"xl/{drTarget.Substring(3)}" : $"xl/worksheets/{drTarget}");
        var drEntry = archive.GetEntry(drPath);
        if (drEntry == null) return;

        string drRelsPath = drPath.Replace("drawings/", "drawings/_rels/") + ".rels";
        var drRelsEntry = archive.GetEntry(drRelsPath);
        if (drRelsEntry == null) return;

        using var drRelsStream = drRelsEntry.Open();
        var drRelsDoc = XDocument.Load(drRelsStream);
        var chartRels = drRelsDoc.Root?.Elements().Where(e => e.Attribute("Type")?.Value.EndsWith("chart") == true);
        if (chartRels == null) return;

        foreach (var cRel in chartRels)
        {
            string cTarget = cRel.Attribute("Target")?.Value ?? string.Empty;
            string cPath = cTarget.StartsWith("/") ? cTarget.Substring(1) : (cTarget.StartsWith("../") ? $"xl/{cTarget.Substring(3)}" : $"xl/drawings/{cTarget}");
            var cEntry = archive.CreateEntry(cPath);
            if (cEntry == null) continue;

            using var cStream = cEntry.Open();
            var cDoc = XDocument.Load(cStream);
            var chartEl = cDoc.Root?.Element(nsChart + "chart");
            if (chartEl != null)
            {
                string title = chartEl.Element(nsChart + "title")?.Element(nsChart + "tx")?.Element(nsChart + "rich")?.Element(nsMain + "p")?.Element(nsMain + "r")?.Element(nsMain + "t")?.Value
                            ?? chartEl.Element(nsChart + "title")?.Element(nsChart + "tx")?.Element(nsChart + "v")?.Value
                            ?? "Chart";

                var plotArea = chartEl.Element(nsChart + "plotArea");
                if (plotArea != null)
                {
                    var chartTypeEl = plotArea.Elements().FirstOrDefault(e => e.Name.Namespace == nsChart && e.Name.LocalName.EndsWith("Chart"));
                    string tag = chartTypeEl?.Name.LocalName ?? "barChart";
                    var cType = tag switch
                    {
                        "lineChart" => Drawings.ChartType.Line,
                        "pieChart" => Drawings.ChartType.Pie,
                        "doughnutChart" => Drawings.ChartType.Doughnut,
                        "areaChart" => Drawings.ChartType.Area,
                        "scatterChart" => Drawings.ChartType.Scatter,
                        _ => Drawings.ChartType.Column
                    };

                    var chartObj = new Drawings.ExcelChart(cType, topRow: 5, leftColumn: 1)
                    {
                        Title = title
                    };

                    if (chartTypeEl != null)
                    {
                        foreach (var serEl in chartTypeEl.Elements(nsChart + "ser"))
                        {
                            string serName = serEl.Element(nsChart + "tx")?.Element(nsChart + "v")?.Value ?? "Series";
                            string valRef = serEl.Element(nsChart + "val")?.Element(nsChart + "numRef")?.Element(nsChart + "f")?.Value ?? string.Empty;
                            string? catRef = serEl.Element(nsChart + "cat")?.Element(nsChart + "strRef")?.Element(nsChart + "f")?.Value;
                            if (!string.IsNullOrEmpty(catRef)) chartObj.CategoryRange = catRef;

                            if (!string.IsNullOrEmpty(valRef))
                            {
                                chartObj.AddSeries(serName, valRef);
                            }
                        }
                    }
                    worksheet.Charts.Add(chartObj);
                }
            }
        }
    }

    private static void WriteStylesXml(Stream stream, StyleRegistry registry)
    {
        var styles = registry.RegisteredStyles;

        // 1. Unique Number Formats (custom numFmtId starts at 164)
        var customNumFmts = new List<(int Id, string Code)>();
        var numFmtMap = new Dictionary<string, int>();

        // Standard built-in number formats
        numFmtMap["General"] = 0;
        numFmtMap["0"] = 1;
        numFmtMap["0.00"] = 2;
        numFmtMap["#,##0"] = 3;
        numFmtMap["#,##0.00"] = 4;
        numFmtMap["yyyy-mm-dd"] = 14;

        int nextNumFmtId = 164;
        foreach (var style in styles)
        {
            string code = style.NumberFormat.FormatCode ?? "General";
            if (!numFmtMap.ContainsKey(code))
            {
                int id = nextNumFmtId++;
                numFmtMap[code] = id;
                customNumFmts.Add((id, code));
            }
        }

        // 2. Unique Fonts (index 0 is default Calibri 11pt)
        var fonts = new List<ExcelFont> { new ExcelFont("Calibri", 11) };
        var fontMap = new Dictionary<ExcelFont, int> { [fonts[0]] = 0 };
        foreach (var style in styles)
        {
            var f = style.Font;
            if (string.IsNullOrEmpty(f.Name)) f = new ExcelFont("Calibri", f.Size == 0 ? 11 : f.Size, f.Bold, f.Italic, f.Underline, f.Color);
            if (!fontMap.ContainsKey(f))
            {
                fontMap[f] = fonts.Count;
                fonts.Add(f);
            }
        }

        // 3. Unique Fills (index 0 is none, index 1 is gray125)
        var fills = new List<ExcelFill>
        {
            new ExcelFill(ExcelFillPatternType.None),
            new ExcelFill(ExcelFillPatternType.Gray125)
        };
        var fillMap = new Dictionary<ExcelFill, int>
        {
            [fills[0]] = 0,
            [fills[1]] = 1
        };
        foreach (var style in styles)
        {
            var fill = style.Fill;
            if (!fillMap.ContainsKey(fill))
            {
                fillMap[fill] = fills.Count;
                fills.Add(fill);
            }
        }

        // 4. Unique Borders (index 0 is empty border)
        var borders = new List<ExcelBorder> { default };
        var borderMap = new Dictionary<ExcelBorder, int> { [default] = 0 };
        foreach (var style in styles)
        {
            var border = style.Border;
            if (!borderMap.ContainsKey(border))
            {
                borderMap[border] = borders.Count;
                borders.Add(border);
            }
        }

        // Build XML
        var doc = new XDocument(
            new XElement(nsSpreadsheet + "styleSheet",
                new XAttribute("xmlns", nsSpreadsheet.NamespaceName),

                // Number Formats
                customNumFmts.Count > 0 ? new XElement(nsSpreadsheet + "numFmts",
                    new XAttribute("count", customNumFmts.Count),
                    customNumFmts.Select(nf => new XElement(nsSpreadsheet + "numFmt",
                        new XAttribute("numFmtId", nf.Id),
                        new XAttribute("formatCode", nf.Code)
                    ))
                ) : null,

                // Fonts
                new XElement(nsSpreadsheet + "fonts", new XAttribute("count", fonts.Count),
                    fonts.Select(f =>
                    {
                        var fEl = new XElement(nsSpreadsheet + "font");
                        if (f.Bold) fEl.Add(new XElement(nsSpreadsheet + "b"));
                        if (f.Italic) fEl.Add(new XElement(nsSpreadsheet + "i"));
                        if (f.Underline) fEl.Add(new XElement(nsSpreadsheet + "u"));
                        fEl.Add(new XElement(nsSpreadsheet + "sz", new XAttribute("val", f.Size > 0 ? f.Size : 11)));
                        if (!string.IsNullOrEmpty(f.Color))
                        {
                            fEl.Add(new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(f.Color!))));
                        }
                        fEl.Add(new XElement(nsSpreadsheet + "name", new XAttribute("val", string.IsNullOrEmpty(f.Name) ? "Calibri" : f.Name)));
                        return fEl;
                    })
                ),

                // Fills
                new XElement(nsSpreadsheet + "fills", new XAttribute("count", fills.Count),
                    fills.Select(fill =>
                    {
                        string pType = fill.PatternType switch
                        {
                            ExcelFillPatternType.Solid => "solid",
                            ExcelFillPatternType.Gray125 => "gray125",
                            _ => "none"
                        };

                        var pEl = new XElement(nsSpreadsheet + "patternFill", new XAttribute("patternType", pType));
                        if (!string.IsNullOrEmpty(fill.BackgroundColor))
                        {
                            pEl.Add(new XElement(nsSpreadsheet + "fgColor", new XAttribute("rgb", FormatArgbHex(fill.BackgroundColor!))));
                        }
                        return new XElement(nsSpreadsheet + "fill", pEl);
                    })
                ),

                // Borders
                new XElement(nsSpreadsheet + "borders", new XAttribute("count", borders.Count),
                    borders.Select(b => new XElement(nsSpreadsheet + "border",
                        CreateBorderItemEl("left", b.Left),
                        CreateBorderItemEl("right", b.Right),
                        CreateBorderItemEl("top", b.Top),
                        CreateBorderItemEl("bottom", b.Bottom)
                    ))
                ),

                // Cell Style Xfs (Default)
                new XElement(nsSpreadsheet + "cellStyleXfs", new XAttribute("count", 1),
                    new XElement(nsSpreadsheet + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0))
                ),

                // Cell Xfs
                new XElement(nsSpreadsheet + "cellXfs", new XAttribute("count", styles.Count),
                    styles.Select(s =>
                    {
                        var fontKey = string.IsNullOrEmpty(s.Font.Name) ? new ExcelFont("Calibri", s.Font.Size == 0 ? 11 : s.Font.Size, s.Font.Bold, s.Font.Italic, s.Font.Underline, s.Font.Color) : s.Font;
                        int numFmtId = numFmtMap.TryGetValue(s.NumberFormat.FormatCode ?? "General", out int nId) ? nId : 0;
                        int fontId = fontMap.TryGetValue(fontKey, out int fId) ? fId : 0;
                        int fillId = fillMap.TryGetValue(s.Fill, out int flId) ? flId : 0;
                        int borderId = borderMap.TryGetValue(s.Border, out int bId) ? bId : 0;

                        var xfEl = new XElement(nsSpreadsheet + "xf",
                            new XAttribute("numFmtId", numFmtId),
                            new XAttribute("fontId", fontId),
                            new XAttribute("fillId", fillId),
                            new XAttribute("borderId", borderId),
                            new XAttribute("xfId", 0),
                            new XAttribute("applyFont", fontId > 0 ? "1" : "0"),
                            new XAttribute("applyFill", fillId > 0 ? "1" : "0"),
                            new XAttribute("applyBorder", borderId > 0 ? "1" : "0"),
                            new XAttribute("applyNumberFormat", numFmtId > 0 ? "1" : "0")
                        );

                        if (s.Alignment.Horizontal != ExcelHorizontalAlignment.General || s.Alignment.Vertical != ExcelVerticalAlignment.Bottom || s.Alignment.WrapText)
                        {
                            string hAlign = s.Alignment.Horizontal.ToString().ToLowerInvariant();
                            string vAlign = s.Alignment.Vertical.ToString().ToLowerInvariant();
                            xfEl.Add(new XElement(nsSpreadsheet + "alignment",
                                new XAttribute("horizontal", hAlign),
                                new XAttribute("vertical", vAlign),
                                s.Alignment.WrapText ? new XAttribute("wrapText", "1") : null
                            ));
                            xfEl.Add(new XAttribute("applyAlignment", "1"));
                        }

                        return xfEl;
                    })
                )
            )
        );
        doc.Save(stream);
    }

    private static XElement CreateBorderItemEl(string name, ExcelBorderItem item)
    {
        if (item.Style == ExcelBorderStyle.None) return new XElement(nsSpreadsheet + name);
        string styleStr = item.Style.ToString().ToLowerInvariant();
        var el = new XElement(nsSpreadsheet + name, new XAttribute("style", styleStr));
        if (!string.IsNullOrEmpty(item.Color))
        {
            el.Add(new XElement(nsSpreadsheet + "color", new XAttribute("rgb", FormatArgbHex(item.Color!))));
        }
        return el;
    }

    private static string FormatValue(object val)
    {
        if (val is double d) return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (val is float f) return f.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (val is decimal dec) return dec.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (val is int i) return i.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (val is long l) return l.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (val is DateTime dt) return dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        return val.ToString() ?? string.Empty;
    }

    private static string FormatArgbHex(string hex)
    {
        if (string.IsNullOrWhiteSpace(hex)) return "FF000000";
        hex = hex.Trim('#');
        if (hex.Length == 6) return "FF" + hex.ToUpperInvariant();
        if (hex.Length == 8) return hex.ToUpperInvariant();
        return "FF" + hex.PadLeft(6, '0').ToUpperInvariant();
    }
}
