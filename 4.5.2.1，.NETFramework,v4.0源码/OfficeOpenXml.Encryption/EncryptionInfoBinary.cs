using System;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Encryption;

internal class EncryptionInfoBinary : EncryptionInfo
{
	internal Flags Flags;

	internal uint HeaderSize;

	internal EncryptionHeader Header;

	internal EncryptionVerifier Verifier;

	internal override void Read(byte[] data)
	{
		Flags = (Flags)BitConverter.ToInt32(data, 4);
		HeaderSize = (uint)BitConverter.ToInt32(data, 8);
		Header = new EncryptionHeader();
		Header.Flags = (Flags)BitConverter.ToInt32(data, 12);
		Header.SizeExtra = BitConverter.ToInt32(data, 16);
		Header.AlgID = (AlgorithmID)BitConverter.ToInt32(data, 20);
		Header.AlgIDHash = (AlgorithmHashID)BitConverter.ToInt32(data, 24);
		Header.KeySize = BitConverter.ToInt32(data, 28);
		Header.ProviderType = (ProviderType)BitConverter.ToInt32(data, 32);
		Header.Reserved1 = BitConverter.ToInt32(data, 36);
		Header.Reserved2 = BitConverter.ToInt32(data, 40);
		byte[] array = new byte[HeaderSize - 34];
		Array.Copy(data, 44, array, 0, (int)(HeaderSize - 34));
		Header.CSPName = Encoding.Unicode.GetString(array);
		int num = (int)(HeaderSize + 12);
		Verifier = new EncryptionVerifier();
		Verifier.SaltSize = (uint)BitConverter.ToInt32(data, num);
		Verifier.Salt = new byte[Verifier.SaltSize];
		Array.Copy(data, num + 4, Verifier.Salt, 0, (int)Verifier.SaltSize);
		Verifier.EncryptedVerifier = new byte[16];
		Array.Copy(data, num + 20, Verifier.EncryptedVerifier, 0, 16);
		Verifier.VerifierHashSize = (uint)BitConverter.ToInt32(data, num + 36);
		Verifier.EncryptedVerifierHash = new byte[Verifier.VerifierHashSize];
		Array.Copy(data, num + 40, Verifier.EncryptedVerifierHash, 0, (int)Verifier.VerifierHashSize);
	}

	internal byte[] WriteBinary()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write(MajorVersion);
		binaryWriter.Write(MinorVersion);
		binaryWriter.Write((int)Flags);
		byte[] array = Header.WriteBinary();
		binaryWriter.Write((uint)array.Length);
		binaryWriter.Write(array);
		binaryWriter.Write(Verifier.WriteBinary());
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}
}
