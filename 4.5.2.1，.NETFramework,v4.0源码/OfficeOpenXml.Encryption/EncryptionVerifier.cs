using System.IO;

namespace OfficeOpenXml.Encryption;

internal class EncryptionVerifier
{
	internal uint SaltSize;

	internal byte[] Salt;

	internal byte[] EncryptedVerifier;

	internal uint VerifierHashSize;

	internal byte[] EncryptedVerifierHash;

	internal byte[] WriteBinary()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write(SaltSize);
		binaryWriter.Write(Salt);
		binaryWriter.Write(EncryptedVerifier);
		binaryWriter.Write(20);
		binaryWriter.Write(EncryptedVerifierHash);
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}
}
