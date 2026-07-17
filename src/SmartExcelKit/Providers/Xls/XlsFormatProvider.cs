using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Providers.Xls;

/// <summary>
/// A high-performance, lightweight legacy Excel XLS (BIFF8) format provider.
/// Resolves compound directory structures and maps binary records directly to Workbook sheets.
/// </summary>
internal sealed class XlsFormatProvider : IWorkbookFormatProvider
{
    internal sealed class BiffRecord
    {
        public ushort Code { get; }
        public byte[] Data { get; }
        public int FileOffset { get; }

        public BiffRecord(ushort code, byte[] data, int fileOffset)
        {
            Code = code;
            Data = data;
            FileOffset = fileOffset;
        }
    }

    /// <inheritdoc />
    public void Read(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            // 1. Copy stream to byte array to allow fast in-memory compound directory seek
            byte[] fileBytes;
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            var oleDoc = new Ole2Document(fileBytes);
            byte[] workbookBytes;

            try
            {
                workbookBytes = oleDoc.GetStream("Workbook");
            }
            catch (FileNotFoundException)
            {
                // Fallback for older formats
                workbookBytes = oleDoc.GetStream("Book");
            }

            // 2. Parse raw records and concatenate CONTINUE records automatically
            var records = ReadRecords(workbookBytes, out var offsetToRecordIndex);

            var boundsheets = new List<(string Name, int FileOffset)>();
            var sharedStrings = new List<string>();

            // 3. First Pass: Read BOUNDSHEETS and Shared String Table (SST)
            foreach (var rec in records)
            {
                if (rec.Code == 0x0085) // BOUNDSHEET
                {
                    uint streamPos = BitConverter.ToUInt32(rec.Data, 0);
                    byte nameLen = rec.Data[6];
                    byte options = rec.Data[7];
                    bool isUnicode = (options & 0x01) != 0;

                    string name = isUnicode
                        ? System.Text.Encoding.Unicode.GetString(rec.Data, 8, nameLen * 2)
                        : System.Text.Encoding.ASCII.GetString(rec.Data, 8, nameLen);

                    boundsheets.Add((name, (int)streamPos));
                }
                else if (rec.Code == 0x00FC) // SST (Shared String Table)
                {
                    int uniqueStrings = BitConverter.ToInt32(rec.Data, 4);
                    int sstOffset = 8;
                    for (int sIdx = 0; sIdx < uniqueStrings; sIdx++)
                    {
                        var (val, bytesRead) = ParseBiff8String(rec.Data, sstOffset);
                        if (bytesRead == 0) break;
                        sharedStrings.Add(val);
                        sstOffset += bytesRead;
                    }
                }
            }

            // Clear existing worksheets
            while (workbook.Worksheets.Count > 0)
            {
                workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
            }

            // 4. Second Pass: Read cell values for each sheet
            foreach (var (sheetName, fileOffset) in boundsheets)
            {
                if (!offsetToRecordIndex.TryGetValue(fileOffset, out int recordIndex)) continue;

                var sheet = workbook.AddWorksheet(sheetName);

                // Collect worksheet records sequentially
                var sheetRecords = new List<BiffRecord>();
                int rIdx = recordIndex;

                // Add first BOF
                sheetRecords.Add(records[rIdx]);
                rIdx++;

                while (rIdx < records.Count)
                {
                    var rec = records[rIdx];
                    if (rec.Code == 0x000A) // EOF of worksheet
                    {
                        break;
                    }
                    sheetRecords.Add(rec);
                    rIdx++;
                }

                // Process sheet records
                for (rIdx = 0; rIdx < sheetRecords.Count; rIdx++)
                {
                    var rec = sheetRecords[rIdx];

                    if (rec.Code == 0x0203) // NUMBER
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort col = BitConverter.ToUInt16(rec.Data, 2);
                        double val = BitConverter.ToDouble(rec.Data, 6);
                        sheet.Cell(row + 1, col + 1).Value = val;
                    }
                    else if (rec.Code == 0x027E) // RK
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort col = BitConverter.ToUInt16(rec.Data, 2);
                        int rkVal = BitConverter.ToInt32(rec.Data, 6);
                        sheet.Cell(row + 1, col + 1).Value = DecodeRkValue(rkVal);
                    }
                    else if (rec.Code == 0x00BD) // MULRK
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort startCol = BitConverter.ToUInt16(rec.Data, 2);
                        ushort endCol = BitConverter.ToUInt16(rec.Data, rec.Data.Length - 2);

                        int numCells = endCol - startCol + 1;
                        for (int c = 0; c < numCells; c++)
                        {
                            int rkVal = BitConverter.ToInt32(rec.Data, 6 + c * 6);
                            sheet.Cell(row + 1, startCol + c + 1).Value = DecodeRkValue(rkVal);
                        }
                    }
                    else if (rec.Code == 0x00FD) // LABELSST
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort col = BitConverter.ToUInt16(rec.Data, 2);
                        int sstIndex = BitConverter.ToInt32(rec.Data, 6);
                        if (sstIndex >= 0 && sstIndex < sharedStrings.Count)
                        {
                            sheet.Cell(row + 1, col + 1).Value = sharedStrings[sstIndex];
                        }
                    }
                    else if (rec.Code == 0x0204) // LABEL
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort col = BitConverter.ToUInt16(rec.Data, 2);
                        var (val, _) = ParseBiff8String(rec.Data, 6);
                        sheet.Cell(row + 1, col + 1).Value = val;
                    }
                    else if (rec.Code == 0x0205) // BOOLERR
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort col = BitConverter.ToUInt16(rec.Data, 2);
                        byte val = rec.Data[6];
                        byte isError = rec.Data[7];
                        if (isError == 0)
                        {
                            sheet.Cell(row + 1, col + 1).Value = val != 0;
                        }
                    }
                    else if (rec.Code == 0x0006) // FORMULA
                    {
                        ushort row = BitConverter.ToUInt16(rec.Data, 0);
                        ushort col = BitConverter.ToUInt16(rec.Data, 2);

                        bool isNum = true;
                        if (rec.Data[6] == 0x00 && rec.Data[12] == 0xFF && rec.Data[13] == 0xFF)
                        {
                            // Evaluated string value is stored in the next STRING record
                            isNum = false;
                            if (rIdx + 1 < sheetRecords.Count && sheetRecords[rIdx + 1].Code == 0x0207) // STRING record
                            {
                                var strRec = sheetRecords[rIdx + 1];
                                var (s, _) = ParseBiff8String(strRec.Data, 0);
                                sheet.Cell(row + 1, col + 1).Value = s;
                            }
                        }
                        else if (rec.Data[6] == 0x01 && rec.Data[12] == 0xFF && rec.Data[13] == 0xFF)
                        {
                            isNum = false;
                            sheet.Cell(row + 1, col + 1).Value = rec.Data[8] != 0;
                        }
                        else if (rec.Data[6] == 0x02 && rec.Data[12] == 0xFF && rec.Data[13] == 0xFF)
                        {
                            isNum = false; // Evaluated error code
                        }

                        if (isNum)
                        {
                            double dVal = BitConverter.ToDouble(rec.Data, 6);
                            sheet.Cell(row + 1, col + 1).Value = dVal;
                        }
                    }
                }
            }
        }
        catch (Exception ex) when (ex is not SmartExcelException)
        {
            throw new ParsingException("Failed to read legacy XLS binary file.", "XLS_READ_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public void Write(Stream stream, ExcelWorkbook workbook)
    {
        throw new NotSupportedException("Writing legacy XLS binary format is not supported. Please export workbooks to XLSX instead.");
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
        throw new NotSupportedException("Writing legacy XLS binary format is not supported. Please export workbooks to XLSX instead.");
    }

    private static List<BiffRecord> ReadRecords(byte[] workbookData, out Dictionary<int, int> offsetToRecordIndex)
    {
        var records = new List<BiffRecord>();
        offsetToRecordIndex = new Dictionary<int, int>();
        int offset = 0;

        while (offset + 4 <= workbookData.Length)
        {
            int currentOffset = offset;
            ushort code = BitConverter.ToUInt16(workbookData, offset);
            ushort len = BitConverter.ToUInt16(workbookData, offset + 2);
            offset += 4;

            byte[] data = new byte[len];
            if (offset + len <= workbookData.Length)
            {
                Array.Copy(workbookData, offset, data, 0, len);
            }
            offset += len;

            if (code == 0x003C && records.Count > 0) // CONTINUE record
            {
                var prev = records[records.Count - 1];
                byte[] merged = new byte[prev.Data.Length + data.Length];
                Array.Copy(prev.Data, 0, merged, 0, prev.Data.Length);
                Array.Copy(data, 0, merged, prev.Data.Length, data.Length);

                records[records.Count - 1] = new BiffRecord(prev.Code, merged, prev.FileOffset);
            }
            else
            {
                offsetToRecordIndex[currentOffset] = records.Count;
                records.Add(new BiffRecord(code, data, currentOffset));
            }
        }
        return records;
    }

    private static (string Value, int BytesRead) ParseBiff8String(byte[] data, int startOffset)
    {
        if (startOffset + 3 > data.Length) return (string.Empty, 0);

        ushort charCount = BitConverter.ToUInt16(data, startOffset);
        byte options = data[startOffset + 2];
        int offset = startOffset + 3;

        bool isUnicode = (options & 0x01) != 0;
        bool hasRichText = (options & 0x08) != 0;
        bool hasPhonetic = (options & 0x04) != 0;

        int runsCount = 0;
        if (hasRichText)
        {
            if (offset + 2 > data.Length) return (string.Empty, 0);
            runsCount = BitConverter.ToUInt16(data, offset);
            offset += 2;
        }

        int phoneticSize = 0;
        if (hasPhonetic)
        {
            if (offset + 4 > data.Length) return (string.Empty, 0);
            phoneticSize = BitConverter.ToInt32(data, offset);
            offset += 4;
        }

        string val;
        int stringBytes = isUnicode ? charCount * 2 : charCount;
        if (offset + stringBytes > data.Length) stringBytes = data.Length - offset;

        if (stringBytes > 0)
        {
            val = isUnicode
                ? System.Text.Encoding.Unicode.GetString(data, offset, stringBytes)
                : System.Text.Encoding.ASCII.GetString(data, offset, stringBytes);
            offset += stringBytes;
        }
        else
        {
            val = string.Empty;
        }

        offset += runsCount * 4;
        offset += phoneticSize;

        return (val, offset - startOffset);
    }

    private static double DecodeRkValue(int rk)
    {
        double val;
        if ((rk & 0x02) != 0) // Integer
        {
            val = rk >> 2;
        }
        else // Double
        {
            // High 30 bits of double (low 32 bits are zero)
            long doubleBits = ((long)rk & 0xFFFFFFFC) << 32;
            val = BitConverter.ToDouble(BitConverter.GetBytes(doubleBits), 0);
        }

        if ((rk & 0x01) == 0) // Scaled by 100
        {
            val /= 100.0;
        }
        return val;
    }
}
