using System;

namespace OfficeOpenXml.Encryption;

internal abstract class EncryptionInfo
{
	internal short MajorVersion;

	internal short MinorVersion;

	internal abstract void Read(byte[] data);

	internal static EncryptionInfo ReadBinary(byte[] data)
	{
		short num = BitConverter.ToInt16(data, 0);
		short num2 = BitConverter.ToInt16(data, 2);
		EncryptionInfo encryptionInfo;
		if ((num2 == 2 || num2 == 3) && num <= 4)
		{
			encryptionInfo = new EncryptionInfoBinary();
		}
		else
		{
			if (num != 4 || num2 != 4)
			{
				throw new NotSupportedException("Unsupported encryption format");
			}
			encryptionInfo = new EncryptionInfoAgile();
		}
		encryptionInfo.MajorVersion = num;
		encryptionInfo.MinorVersion = num2;
		encryptionInfo.Read(data);
		return encryptionInfo;
	}
}
