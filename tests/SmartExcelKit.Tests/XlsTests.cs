using FluentAssertions;
using Xunit;

namespace SmartExcelKit.Tests;

public class XlsTests
{
    [Fact]
    public void XlsProvider_ShouldReadBiff8FormatSymmetrically()
    {
        // Assemble BIFF8 Workbook records in memory
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms);

        // 1. Workbook BOF (0x0809)
        bw.Write((ushort)0x0809);
        bw.Write((ushort)8);
        bw.Write(new byte[] { 0x00, 0x06, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00 });

        // 2. BOUNDSHEET (0x0085) - pointing to offset 54 (0x36)
        bw.Write((ushort)0x0085);
        bw.Write((ushort)14);
        bw.Write((uint)0x36); // Sheet BOF offset
        bw.Write((byte)0x00); // visibility
        bw.Write((byte)0x00); // type
        bw.Write((byte)6);    // name length
        bw.Write((byte)0x00); // ASCII option
        bw.Write(System.Text.Encoding.ASCII.GetBytes("Sheet1"));

        // 3. SST (0x00FC)
        bw.Write((ushort)0x00FC);
        bw.Write((ushort)16);
        bw.Write((uint)1);    // total strings
        bw.Write((uint)1);    // unique strings
        bw.Write((ushort)5);  // character count
        bw.Write((byte)0x00); // option (ASCII)
        bw.Write(System.Text.Encoding.ASCII.GetBytes("Alice"));

        // 4. EOF (0x000A)
        bw.Write((ushort)0x000A);
        bw.Write((ushort)0);

        // 5. Worksheet BOF (0x0809) - located at offset 54
        bw.Write((ushort)0x0809);
        bw.Write((ushort)8);
        bw.Write(new byte[] { 0x00, 0x06, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00 });

        // 6. NUMBER (0x0203) - cell A1 = 123.45
        bw.Write((ushort)0x0203);
        bw.Write((ushort)14);
        bw.Write((ushort)0);  // row 0
        bw.Write((ushort)0);  // col 0
        bw.Write((ushort)0);  // XF
        bw.Write(123.45);     // value

        // 7. LABELSST (0x00FD) - cell B1 = "Alice"
        bw.Write((ushort)0x00FD);
        bw.Write((ushort)10);
        bw.Write((ushort)0);  // row 0
        bw.Write((ushort)1);  // col 1
        bw.Write((ushort)0);  // XF
        bw.Write((uint)0);    // SST index 0

        // 8. EOF (0x000A)
        bw.Write((ushort)0x000A);
        bw.Write((ushort)0);

        byte[] workbookStream = ms.ToArray();
        byte[] xlsFileBytes = CreateMockXls(workbookStream);

        // Act: Open workbook
        using var readStream = new MemoryStream(xlsFileBytes);
        using var workbook = ExcelWorkbook.Open(readStream);

        // Assert
        workbook.Worksheets.Count.Should().Be(1);
        var sheet = workbook.Worksheets[0];
        sheet.Name.Should().Be("Sheet1");

        // Cell values check
        sheet.Cell("A1").Value.Should().Be(123.45);
        sheet.Cell("B1").Value.Should().Be("Alice");
    }

    private static byte[] CreateMockXls(byte[] workbookStream)
    {
        byte[] file = new byte[512 * 4]; // Header, Dir, FAT, Workbook
        // Magic
        byte[] magic = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        Array.Copy(magic, 0, file, 0, 8);

        // Sector shift (9 -> 512 bytes)
        file[30] = 9;
        file[31] = 0;

        // Mini sector shift (6 -> 64 bytes)
        file[32] = 6;
        file[33] = 0;

        // Number of FAT sectors = 1
        BitConverter.GetBytes(1).CopyTo(file, 44);

        // First Directory sector ID = 0 (sector 0)
        BitConverter.GetBytes(0).CopyTo(file, 48);

        // First Mini-FAT sector ID = -2 (End of chain)
        BitConverter.GetBytes(-2).CopyTo(file, 60);

        // FAT table links in header start at offset 76
        for (int i = 0; i < 109; i++)
        {
            BitConverter.GetBytes(-1).CopyTo(file, 76 + i * 4);
        }
        // First FAT sector ID = 1 (sector 1)
        BitConverter.GetBytes(1).CopyTo(file, 76);

        // Sector 0: Directory Entries
        int rootOffset = 512;
        byte[] rootName = System.Text.Encoding.Unicode.GetBytes("Root Entry\0");
        Array.Copy(rootName, 0, file, rootOffset, rootName.Length);
        BitConverter.GetBytes((short)rootName.Length).CopyTo(file, rootOffset + 64);
        file[rootOffset + 66] = 5; // Root Storage
        BitConverter.GetBytes(-2).CopyTo(file, rootOffset + 116);

        int wbOffset = 512 + 128;
        byte[] wbName = System.Text.Encoding.Unicode.GetBytes("Workbook\0");
        Array.Copy(wbName, 0, file, wbOffset, wbName.Length);
        BitConverter.GetBytes((short)wbName.Length).CopyTo(file, wbOffset + 64);
        file[wbOffset + 66] = 2; // Stream
        BitConverter.GetBytes(2).CopyTo(file, wbOffset + 116); // sector 2
        BitConverter.GetBytes(workbookStream.Length).CopyTo(file, wbOffset + 120);

        // Sector 1: FAT
        int fatOffset = 512 * 2;
        BitConverter.GetBytes(-2).CopyTo(file, fatOffset + 0 * 4);
        BitConverter.GetBytes(-3).CopyTo(file, fatOffset + 1 * 4);
        BitConverter.GetBytes(-2).CopyTo(file, fatOffset + 2 * 4); // End of Workbook
        BitConverter.GetBytes(-2).CopyTo(file, fatOffset + 3 * 4);

        // Sector 2: Workbook stream
        int streamOffset = 512 * 3;
        Array.Copy(workbookStream, 0, file, streamOffset, Math.Min(workbookStream.Length, 512));

        return file;
    }
}
