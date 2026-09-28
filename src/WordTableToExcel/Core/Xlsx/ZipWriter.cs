using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace WordTableToExcel.Core.Xlsx
{
    /// <summary>
    /// Écriture minimale d'une archive ZIP (format utilisé par les fichiers .xlsx),
    /// sans dépendance à System.IO.Packaging ni à des bibliothèques tierces.
    /// </summary>
    public sealed class ZipWriter : IDisposable
    {
        private sealed class Entry
        {
            public string Name;
            public uint Crc;
            public int CompressedSize;
            public int UncompressedSize;
            public ushort Method;
            public long Offset;
        }

        private readonly Stream _output;
        private readonly List<Entry> _entries = new List<Entry>();
        private readonly ushort _dosTime;
        private readonly ushort _dosDate;
        private bool _finished;

        public ZipWriter(Stream output)
        {
            if (output == null) throw new ArgumentNullException("output");
            _output = output;
            var now = DateTime.Now;
            _dosTime = (ushort)((now.Hour << 11) | (now.Minute << 5) | (now.Second / 2));
            _dosDate = (ushort)(((Math.Max(1980, now.Year) - 1980) << 9) | (now.Month << 5) | now.Day);
        }

        public void AddEntry(string name, string content)
        {
            AddEntry(name, new UTF8Encoding(false).GetBytes(content));
        }

        public void AddEntry(string name, byte[] data)
        {
            if (_finished) throw new InvalidOperationException("Archive déjà finalisée.");
            byte[] compressed = Deflate(data);
            ushort method = 8;
            if (compressed.Length >= data.Length)
            {
                compressed = data;
                method = 0;
            }

            var entry = new Entry
            {
                Name = name,
                Crc = Crc32.Compute(data),
                CompressedSize = compressed.Length,
                UncompressedSize = data.Length,
                Method = method,
                Offset = _output.Position
            };
            _entries.Add(entry);

            byte[] nameBytes = Encoding.UTF8.GetBytes(name);
            var w = new BinaryWriter(_output);
            w.Write(0x04034b50u);
            w.Write((ushort)20);          // version nécessaire
            w.Write((ushort)0x0800);      // noms en UTF-8
            w.Write(method);
            w.Write(_dosTime);
            w.Write(_dosDate);
            w.Write(entry.Crc);
            w.Write(entry.CompressedSize);
            w.Write(entry.UncompressedSize);
            w.Write((ushort)nameBytes.Length);
            w.Write((ushort)0);
            w.Write(nameBytes);
            w.Write(compressed);
            w.Flush();
        }

        public void Finish()
        {
            if (_finished) return;
            _finished = true;
            var w = new BinaryWriter(_output);
            long centralStart = _output.Position;
            foreach (var e in _entries)
            {
                byte[] nameBytes = Encoding.UTF8.GetBytes(e.Name);
                w.Write(0x02014b50u);
                w.Write((ushort)20);      // version créée par
                w.Write((ushort)20);      // version nécessaire
                w.Write((ushort)0x0800);
                w.Write(e.Method);
                w.Write(_dosTime);
                w.Write(_dosDate);
                w.Write(e.Crc);
                w.Write(e.CompressedSize);
                w.Write(e.UncompressedSize);
                w.Write((ushort)nameBytes.Length);
                w.Write((ushort)0);       // extra
                w.Write((ushort)0);       // commentaire
                w.Write((ushort)0);       // disque
                w.Write((ushort)0);       // attributs internes
                w.Write(0u);              // attributs externes
                w.Write((uint)e.Offset);
                w.Write(nameBytes);
            }
            long centralSize = _output.Position - centralStart;
            w.Write(0x06054b50u);
            w.Write((ushort)0);
            w.Write((ushort)0);
            w.Write((ushort)_entries.Count);
            w.Write((ushort)_entries.Count);
            w.Write((uint)centralSize);
            w.Write((uint)centralStart);
            w.Write((ushort)0);
            w.Flush();
        }

        public void Dispose()
        {
            Finish();
        }

        private static byte[] Deflate(byte[] data)
        {
            using (var ms = new MemoryStream())
            {
                using (var deflate = new DeflateStream(ms, CompressionMode.Compress, true))
                {
                    deflate.Write(data, 0, data.Length);
                }
                return ms.ToArray();
            }
        }
    }

    internal static class Crc32
    {
        private static readonly uint[] Table = BuildTable();

        public static uint Compute(byte[] data)
        {
            uint crc = 0xFFFFFFFFu;
            foreach (byte b in data)
            {
                crc = Table[(crc ^ b) & 0xFF] ^ (crc >> 8);
            }
            return crc ^ 0xFFFFFFFFu;
        }

        private static uint[] BuildTable()
        {
            var table = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int k = 0; k < 8; k++)
                {
                    c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
                }
                table[i] = c;
            }
            return table;
        }
    }
}
