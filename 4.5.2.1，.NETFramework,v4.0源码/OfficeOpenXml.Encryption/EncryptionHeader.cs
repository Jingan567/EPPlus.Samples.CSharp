using System.IO;
using System.Text;

namespace OfficeOpenXml.Encryption;

internal class EncryptionHeader
{
	internal Flags Flags;

	internal int SizeExtra;

	internal AlgorithmID AlgID;

	internal AlgorithmHashID AlgIDHash;

	internal int KeySize;

	internal ProviderType ProviderType;

	internal int Reserved1;

	internal int Reserved2;

	internal string CSPName;

	internal byte[] WriteBinary()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write((int)Flags);
		binaryWriter.Write(SizeExtra);
		binaryWriter.Write((int)AlgID);
		binaryWriter.Write((int)AlgIDHash);
		binaryWriter.Write(KeySize);
		binaryWriter.Write((int)ProviderType);
		binaryWriter.Write(Reserved1);
		binaryWriter.Write(Reserved2);
		binaryWriter.Write(Encoding.Unicode.GetBytes(CSPName));
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}
}
