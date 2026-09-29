using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using WordTableToExcel.Core.Xlsx;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Lecture d'une archive ZIP (format des fichiers .xlsx) à partir de son répertoire central :
    /// entrées stockées ou compressées (deflate), extensions ZIP64, contrôle CRC.
    /// Sans dépendance externe (compatible .NET Framework 4.0).
    /// </summary>
    public sealed class ZipReader
    {
        private const long MaxEntrySize = 512L * 1024 * 1024;

        private sealed class Entry
        {
            public string Name;
            public ushort Method;
            public ushort Flags;
            public uint Crc;
            public long CompressedSize;
            public long UncompressedSize;
            public long LocalHeaderOffset;
        }

        private readonly byte[] _data;
        private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);

        public ZipReader(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            _data = data;
            ReadCentralDirectory();
        }

        public IEnumerable<string> Names
        {
            get { return _entries.Keys; }
        }

        public bool Contains(string name)
        {
            return name != null && _entries.ContainsKey(Normalize(name));
        }

        /// <summary>Contenu décompressé de l'entrée, ou null si elle n'existe pas.</summary>
        public byte[] Read(string name)
        {
            Entry entry;
            if (name == null || !_entries.TryGetValue(Normalize(name), out entry)) return null;
            if ((entry.Flags & 1) != 0) throw new InvalidDataException("L'entrée « " + entry.Name + " » est chiffrée.");
            if (entry.UncompressedSize > MaxEntrySize) throw new InvalidDataException("L'entrée « " + entry.Name + " » est trop volumineuse.");

            long offset = entry.LocalHeaderOffset;
            if (offset < 0 || offset + 30 > _data.Length || ReadUInt32(offset) != 0x04034b50u)
            {
                throw new InvalidDataException("En-tête local invalide pour « " + entry.Name + " ».");
            }
            int nameLength = ReadUInt16(offset + 26);
            int extraLength = ReadUInt16(offset + 28);
            long dataStart = offset + 30 + nameLength + extraLength;
            if (dataStart + entry.CompressedSize > _data.Length) throw new InvalidDataException("Archive tronquée (« " + entry.Name + " »).");

            byte[] result;
            if (entry.Method == 0)
            {
                result = new byte[entry.CompressedSize];
                Buffer.BlockCopy(_data, (int)dataStart, result, 0, result.Length);
            }
            else if (entry.Method == 8)
            {
                using (var input = new MemoryStream(_data, (int)dataStart, (int)entry.CompressedSize, false))
                using (var inflate = new DeflateStream(input, CompressionMode.Decompress))
                using (var output = new MemoryStream(entry.UncompressedSize > 0 && entry.UncompressedSize < int.MaxValue ? (int)entry.UncompressedSize : 4096))
                {
                    var buffer = new byte[81920];
                    int read;
                    while ((read = inflate.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        output.Write(buffer, 0, read);
                        if (output.Length > MaxEntrySize) throw new InvalidDataException("L'entrée « " + entry.Name + " » est trop volumineuse.");
                    }
                    result = output.ToArray();
                }
            }
            else
            {
                throw new InvalidDataException("Méthode de compression non prise en charge (" + entry.Method + ") pour « " + entry.Name + " ».");
            }

            if (result.Length != entry.UncompressedSize || Crc32.Compute(result) != entry.Crc)
            {
                throw new InvalidDataException("Contenu altéré (contrôle CRC) pour « " + entry.Name + " ».");
            }
            return result;
        }

        private void ReadCentralDirectory()
        {
            long eocd = FindEndOfCentralDirectory();
            if (eocd < 0) throw new InvalidDataException("Ce fichier n'est pas une archive ZIP valide.");

            long count = ReadUInt16(eocd + 10);
            long size = ReadUInt32(eocd + 12);
            long offset = ReadUInt32(eocd + 16);

            if (count == 0xFFFF || size == 0xFFFFFFFF || offset == 0xFFFFFFFF)
            {
                // ZIP64 : localisateur juste avant l'enregistrement de fin.
                long locator = eocd - 20;
                if (locator >= 0 && ReadUInt32(locator) == 0x07064b50u)
                {
                    long record = (long)ReadUInt64(locator + 8);
                    if (record >= 0 && record + 56 <= _data.Length && ReadUInt32(record) == 0x06064b50u)
                    {
                        count = (long)ReadUInt64(record + 32);
                        size = (long)ReadUInt64(record + 40);
                        offset = (long)ReadUInt64(record + 48);
                    }
                }
            }

            long position = offset;
            for (long i = 0; i < count; i++)
            {
                if (position + 46 > _data.Length || ReadUInt32(position) != 0x02014b50u)
                {
                    throw new InvalidDataException("Répertoire central de l'archive invalide.");
                }
                var entry = new Entry
                {
                    Flags = ReadUInt16(position + 8),
                    Method = ReadUInt16(position + 10),
                    Crc = ReadUInt32(position + 16),
                    CompressedSize = ReadUInt32(position + 20),
                    UncompressedSize = ReadUInt32(position + 24),
                    LocalHeaderOffset = ReadUInt32(position + 42)
                };
                int nameLength = ReadUInt16(position + 28);
                int extraLength = ReadUInt16(position + 30);
                int commentLength = ReadUInt16(position + 32);
                long nameStart = position + 46;
                if (nameStart + nameLength + extraLength > _data.Length) throw new InvalidDataException("Répertoire central tronqué.");

                entry.Name = Normalize(Encoding.UTF8.GetString(_data, (int)nameStart, nameLength));
                ReadZip64Extra(entry, nameStart + nameLength, extraLength);

                if (!entry.Name.EndsWith("/", StringComparison.Ordinal) && !_entries.ContainsKey(entry.Name))
                {
                    _entries[entry.Name] = entry;
                }
                position = nameStart + nameLength + extraLength + commentLength;
            }
        }

        private void ReadZip64Extra(Entry entry, long start, int length)
        {
            long end = start + length;
            long p = start;
            while (p + 4 <= end)
            {
                int id = ReadUInt16(p);
                int size = ReadUInt16(p + 2);
                long data = p + 4;
                if (id == 0x0001)
                {
                    long q = data;
                    if (entry.UncompressedSize == 0xFFFFFFFF && q + 8 <= data + size) { entry.UncompressedSize = (long)ReadUInt64(q); q += 8; }
                    if (entry.CompressedSize == 0xFFFFFFFF && q + 8 <= data + size) { entry.CompressedSize = (long)ReadUInt64(q); q += 8; }
                    if (entry.LocalHeaderOffset == 0xFFFFFFFF && q + 8 <= data + size) { entry.LocalHeaderOffset = (long)ReadUInt64(q); }
                    return;
                }
                p = data + size;
            }
        }

        private long FindEndOfCentralDirectory()
        {
            long min = Math.Max(0, _data.Length - 22 - 65535);
            for (long p = _data.Length - 22; p >= min; p--)
            {
                if (ReadUInt32(p) == 0x06054b50u) return p;
            }
            return -1;
        }

        private static string Normalize(string name)
        {
            return name.Replace('\\', '/').TrimStart('/');
        }

        private ushort ReadUInt16(long p)
        {
            return (ushort)(_data[p] | (_data[p + 1] << 8));
        }

        private uint ReadUInt32(long p)
        {
            return (uint)(_data[p] | (_data[p + 1] << 8) | (_data[p + 2] << 16) | (_data[p + 3] << 24));
        }

        private ulong ReadUInt64(long p)
        {
            return ReadUInt32(p) | ((ulong)ReadUInt32(p + 4) << 32);
        }
    }
}
