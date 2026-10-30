using System;
using System.IO;
using System.Collections.Generic;

public static class VbaToolkitNative
{
    public static byte[] ReadSectorChain(byte[] bytes, int startSector, int sectorSize, int[] fat)
    {
        var ms = new MemoryStream();
        var visited = new HashSet<int>();
        int s = startSector;
        while (s >= 0 && s != -2 && s != -1 && !visited.Contains(s))
        {
            visited.Add(s);
            int off = (s + 1) * sectorSize;
            if (off + sectorSize > bytes.Length) break;
            ms.Write(bytes, off, sectorSize);
            s = (s < fat.Length) ? fat[s] : -1;
        }
        return ms.ToArray();
    }

    public static byte[] ReadMiniStream(byte[] miniStreamData, int startSector, int size, int miniSectorSize, int[] miniFat)
    {
        var data = new byte[size];
        int s = startSector;
        int written = 0;
        while (s >= 0 && s != -2 && written < size)
        {
            int off = s * miniSectorSize;
            int toRead = Math.Min(miniSectorSize, size - written);
            if (off + toRead <= miniStreamData.Length)
                Array.Copy(miniStreamData, off, data, written, toRead);
            written += miniSectorSize;
            s = (s < miniFat.Length) ? miniFat[s] : -1;
        }
        return data;
    }

    public static byte[] DecompressVba(byte[] data, int offset)
    {
        if (offset >= data.Length || data[offset] != 1) return new byte[0];
        var result = new List<byte>(data.Length * 2);
        int pos = offset + 1;
        while (pos < data.Length - 1)
        {
            if (pos + 1 >= data.Length) break;
            ushort header = BitConverter.ToUInt16(data, pos); pos += 2;
            int chunkSize = (header & 0x0FFF) + 3;
            bool isCompressed = (header & 0x8000) != 0;
            if (!isCompressed)
            {
                int toCopy = Math.Min(chunkSize - 2, data.Length - pos);
                for (int c = 0; c < toCopy; c++) result.Add(data[pos + c]);
                pos += toCopy;
                continue;
            }
            int chunkEnd = pos + chunkSize - 2;
            if (chunkEnd > data.Length) chunkEnd = data.Length;
            int decompStart = result.Count;
            while (pos < chunkEnd)
            {
                if (pos >= data.Length) break;
                byte flagByte = data[pos]; pos++;
                for (int bit = 0; bit < 8 && pos < chunkEnd; bit++)
                {
                    if ((flagByte & (1 << bit)) == 0)
                    {
                        result.Add(data[pos]); pos++;
                    }
                    else
                    {
                        if (pos + 1 >= data.Length) { pos = chunkEnd; break; }
                        ushort token = BitConverter.ToUInt16(data, pos); pos += 2;
                        int dPos = result.Count - decompStart;
                        if (dPos < 1) dPos = 1;
                        int bitCount = 4;
                        while ((1 << bitCount) < dPos) bitCount++;
                        if (bitCount > 12) bitCount = 12;
                        int lengthMask = 0xFFFF >> bitCount;
                        int copyLen = (token & lengthMask) + 3;
                        int copyOff = (token >> (16 - bitCount)) + 1;
                        for (int c = 0; c < copyLen; c++)
                        {
                            int srcIdx = result.Count - copyOff;
                            if (srcIdx >= 0 && srcIdx < result.Count)
                                result.Add(result[srcIdx]);
                        }
                    }
                }
            }
            pos = chunkEnd;
        }
        return result.ToArray();
    }

    public static byte[] CompressVba(byte[] data)
    {
        var result = new MemoryStream();
        result.WriteByte(1);
        int srcPos = 0;
        while (srcPos < data.Length)
        {
            int chunkStart = srcPos;
            int chunkEnd = Math.Min(srcPos + 4096, data.Length);
            var chunkBuf = new MemoryStream();
            int dPos = srcPos;
            while (dPos < chunkEnd)
            {
                long flagPos = chunkBuf.Position;
                chunkBuf.WriteByte(0);
                byte flagByte = 0;
                for (int bit = 0; bit < 8 && dPos < chunkEnd; bit++)
                {
                    int bestLen = 0, bestOff = 0;
                    int decompPos = dPos - chunkStart;
                    if (decompPos < 1) decompPos = 1;
                    int bitCount = 4;
                    while ((1 << bitCount) < decompPos) bitCount++;
                    if (bitCount > 12) bitCount = 12;
                    int maxOff = 1 << bitCount;
                    int maxLen = (0xFFFF >> bitCount) + 3;
                    for (int off = 1; off <= Math.Min(maxOff, dPos - chunkStart); off++)
                    {
                        int matchLen = 0;
                        while (matchLen < maxLen && dPos + matchLen < chunkEnd)
                        {
                            if (data[dPos - off + (matchLen % off)] != data[dPos + matchLen]) break;
                            matchLen++;
                        }
                        if (matchLen >= 3 && matchLen > bestLen) { bestLen = matchLen; bestOff = off; }
                    }
                    if (bestLen >= 3)
                    {
                        flagByte |= (byte)(1 << bit);
                        int token = ((bestOff - 1) << (16 - bitCount)) | (bestLen - 3);
                        chunkBuf.WriteByte((byte)(token & 0xFF));
                        chunkBuf.WriteByte((byte)((token >> 8) & 0xFF));
                        dPos += bestLen;
                    }
                    else
                    {
                        chunkBuf.WriteByte(data[dPos]); dPos++;
                    }
                }
                long savedPos = chunkBuf.Position;
                chunkBuf.Position = flagPos;
                chunkBuf.WriteByte(flagByte);
                chunkBuf.Position = savedPos;
            }
            byte[] compressed = chunkBuf.ToArray();
            srcPos = dPos;
            ushort hdr = (ushort)(0x8000 | 0x3000 | (compressed.Length + 2 - 3));
            result.WriteByte((byte)(hdr & 0xFF));
            result.WriteByte((byte)((hdr >> 8) & 0xFF));
            result.Write(compressed, 0, compressed.Length);
        }
        return result.ToArray();
    }

    public static int FindPattern(byte[] data, byte[] pattern)
    {
        for (int i = 0; i <= data.Length - pattern.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j]) { match = false; break; }
            }
            if (match) return i;
        }
        return -1;
    }
}