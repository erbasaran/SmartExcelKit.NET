namespace SmartExcelKit.Providers.Xls;

/// <summary>
/// A simple parser for the OLE2 Compound File Binary Format (CFBF).
/// Extracts stream payloads (like "Workbook" or "Book") from structured binary sectors.
/// </summary>
internal sealed class Ole2Document
{
    private readonly byte[] _data;
    private readonly int _sectorSize;
    private readonly int _dirStartSector;
    private readonly List<int> _fat = new();

    public Ole2Document(byte[] data)
    {
        _data = data ?? throw new ArgumentNullException(nameof(data));

        if (data.Length < 512 ||
            data[0] != 0xD0 || data[1] != 0xCF || data[2] != 0x11 || data[3] != 0xE0 ||
            data[4] != 0xA1 || data[5] != 0xB1 || data[6] != 0x1A || data[7] != 0xE1)
        {
            throw new InvalidDataException("Invalid OLE2 document magic signature.");
        }

        int sectorShift = BitConverter.ToInt16(data, 30);
        _sectorSize = 1 << sectorShift;

        _dirStartSector = BitConverter.ToInt32(data, 48);

        int numFatSectors = BitConverter.ToInt32(data, 44);
        int firstFatSector = BitConverter.ToInt32(data, 48);

        var fatSectors = new List<int>();

        // Read initial FAT sector links from header
        for (int i = 0; i < 109; i++)
        {
            int sec = BitConverter.ToInt32(data, 76 + i * 4);
            if (sec >= 0) fatSectors.Add(sec);
        }

        // Loop FAT sectors and link references
        foreach (int fatSec in fatSectors)
        {
            byte[] secData = GetSectorData(fatSec);
            for (int offset = 0; offset < _sectorSize; offset += 4)
            {
                _fat.Add(BitConverter.ToInt32(secData, offset));
            }
        }
    }

    private byte[] GetSectorData(int sectorId)
    {
        if (sectorId < 0) throw new ArgumentOutOfRangeException(nameof(sectorId), "Sector ID cannot be negative.");
        int offset = (sectorId + 1) * _sectorSize;
        if (offset + _sectorSize > _data.Length)
        {
            byte[] padded = new byte[_sectorSize];
            int available = _data.Length - offset;
            if (available > 0)
            {
                Array.Copy(_data, offset, padded, 0, available);
            }
            return padded;
        }

        byte[] sector = new byte[_sectorSize];
        Array.Copy(_data, offset, sector, 0, _sectorSize);
        return sector;
    }

    /// <summary>
    /// Extracts a stream payload by name from the OLE2 directory tree.
    /// </summary>
    public byte[] GetStream(string streamName)
    {
        int dirSec = _dirStartSector;
        var dirSectors = new List<int>();
        while (dirSec >= 0 && dirSec < _fat.Count)
        {
            dirSectors.Add(dirSec);
            dirSec = _fat[dirSec];
        }

        foreach (int secId in dirSectors)
        {
            byte[] dirData = GetSectorData(secId);
            for (int entryOffset = 0; entryOffset < _sectorSize; entryOffset += 128)
            {
                int nameLen = BitConverter.ToInt16(dirData, entryOffset + 64);
                if (nameLen > 0 && nameLen <= 64)
                {
                    string name = System.Text.Encoding.Unicode.GetString(dirData, entryOffset, nameLen - 2);
                    if (string.Equals(name, streamName, StringComparison.OrdinalIgnoreCase))
                    {
                        int startSector = BitConverter.ToInt32(dirData, entryOffset + 116);
                        int size = BitConverter.ToInt32(dirData, entryOffset + 120);

                        return ReadStreamSectors(startSector, size);
                    }
                }
            }
        }
        throw new FileNotFoundException($"OLE2 stream '{streamName}' not found in file directory.");
    }

    private byte[] ReadStreamSectors(int startSector, int size)
    {
        using var ms = new MemoryStream();
        int sec = startSector;
        while (sec >= 0 && sec < _fat.Count && ms.Length < size)
        {
            byte[] data = GetSectorData(sec);
            ms.Write(data, 0, data.Length);
            sec = _fat[sec];
        }
        byte[] result = new byte[size];
        Array.Copy(ms.ToArray(), 0, result, 0, size);
        return result;
    }
}
