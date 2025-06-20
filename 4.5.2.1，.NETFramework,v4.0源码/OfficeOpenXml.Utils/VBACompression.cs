using System;
using System.IO;

namespace OfficeOpenXml.Utils;

internal static class VBACompression
{
	internal static byte[] CompressPart(byte[] part)
	{
		MemoryStream memoryStream = new MemoryStream(4096);
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write((byte)1);
		int num = 1;
		int num2 = 4098;
		int startPos = 0;
		int num3 = ((part.Length < 4096) ? part.Length : 4096);
		while (startPos < num3 && num < num2)
		{
			byte[] array = CompressChunk(part, ref startPos);
			if (array == null || array.Length == 0)
			{
				ushort num4 = 5632;
			}
			else
			{
				ushort num4 = (ushort)((uint)(array.Length - 1) & 0xFFFu);
				num4 = (ushort)(num4 | 0xB000u);
				binaryWriter.Write(num4);
				binaryWriter.Write(array);
			}
			num3 = ((part.Length < startPos + 4096) ? part.Length : (startPos + 4096));
		}
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private static byte[] CompressChunk(byte[] buffer, ref int startPos)
	{
		byte[] array = new byte[4096];
		int num = 0;
		int num2 = 1;
		int num3 = startPos;
		int num4 = ((startPos + 4096 < buffer.Length) ? (startPos + 4096) : buffer.Length);
		while (num3 < num4)
		{
			byte b = 0;
			for (int i = 0; i < 8; i++)
			{
				if (num3 - startPos > 0)
				{
					int num5 = -1;
					int num6 = 0;
					int num7 = num3 - 1;
					int lengthBits = GetLengthBits(num3 - startPos);
					int num8 = 16 - lengthBits;
					ushort num9 = (ushort)(65535 >> num8);
					while (num7 >= startPos)
					{
						if (buffer[num7] == buffer[num3])
						{
							int j;
							for (j = 1; buffer.Length > num3 + j && buffer[num7 + j] == buffer[num3 + j] && j < num9 && num3 + j < num4; j++)
							{
							}
							if (j > num6)
							{
								num5 = num7;
								num6 = j;
								if (num6 == num9)
								{
									break;
								}
							}
						}
						num7--;
					}
					if (num6 >= 3)
					{
						b |= (byte)(1 << i);
						Array.Copy(BitConverter.GetBytes((ushort)(((ushort)(num3 - (num5 + 1)) << lengthBits) | (ushort)(num6 - 3))), 0, array, num2, 2);
						num3 += num6;
						num2 += 2;
					}
					else
					{
						array[num2++] = buffer[num3++];
					}
				}
				else
				{
					array[num2++] = buffer[num3++];
				}
				if (num3 >= num4)
				{
					break;
				}
			}
			array[num] = b;
			num = num2++;
		}
		byte[] array2 = new byte[num2 - 1];
		Array.Copy(array, array2, array2.Length);
		startPos = num4;
		return array2;
	}

	internal static byte[] DecompressPart(byte[] part)
	{
		return DecompressPart(part, 0);
	}

	internal static byte[] DecompressPart(byte[] part, int startPos)
	{
		if (part[startPos] != 1)
		{
			return null;
		}
		MemoryStream memoryStream = new MemoryStream(4096);
		int pos = startPos + 1;
		while (pos < part.Length - 1)
		{
			DecompressChunk(memoryStream, part, ref pos);
		}
		return memoryStream.ToArray();
	}

	private static void DecompressChunk(MemoryStream ms, byte[] compBuffer, ref int pos)
	{
		ushort num = BitConverter.ToUInt16(compBuffer, pos);
		int num2 = 0;
		byte[] array = new byte[4198];
		int num3 = (num & 0xFFF) + 3;
		int num4 = pos + num3;
		int num5 = (num & 0x8000) >> 15;
		pos += 2;
		if (num5 == 1)
		{
			while (pos < compBuffer.Length && pos < num4)
			{
				byte b = compBuffer[pos++];
				if (pos >= num4)
				{
					break;
				}
				for (int i = 0; i < 8; i++)
				{
					if ((b & (1 << i)) == 0)
					{
						ms.WriteByte(compBuffer[pos]);
						array[num2++] = compBuffer[pos++];
					}
					else
					{
						ushort num6 = BitConverter.ToUInt16(compBuffer, pos);
						int lengthBits = GetLengthBits(num2);
						int num7 = 16 - lengthBits;
						ushort num8 = (ushort)(65535 >> num7);
						ushort num9 = (ushort)(~num8);
						int num10 = (num8 & num6) + 3;
						int num11 = (num9 & num6) >> lengthBits;
						int num12 = num2 - num11 - 1;
						if (num2 + num10 >= array.Length)
						{
							byte[] array2 = new byte[array.Length + 4098];
							Array.Copy(array, array2, num2);
							array = array2;
						}
						for (int j = 0; j < num10; j++)
						{
							ms.WriteByte(array[num12]);
							array[num2++] = array[num12++];
						}
						pos += 2;
					}
					if (pos >= num4)
					{
						break;
					}
				}
			}
		}
		else
		{
			ms.Write(compBuffer, pos, num3);
			pos += num3;
		}
	}

	private static int GetLengthBits(int decompPos)
	{
		if (decompPos <= 16)
		{
			return 12;
		}
		if (decompPos <= 32)
		{
			return 11;
		}
		if (decompPos <= 64)
		{
			return 10;
		}
		if (decompPos <= 128)
		{
			return 9;
		}
		if (decompPos <= 256)
		{
			return 8;
		}
		if (decompPos <= 512)
		{
			return 7;
		}
		if (decompPos <= 1024)
		{
			return 6;
		}
		if (decompPos <= 2048)
		{
			return 5;
		}
		if (decompPos <= 4096)
		{
			return 4;
		}
		return 12 - (int)Math.Truncate(Math.Log(decompPos - 1 >> 4, 2.0) + 1.0);
	}
}
