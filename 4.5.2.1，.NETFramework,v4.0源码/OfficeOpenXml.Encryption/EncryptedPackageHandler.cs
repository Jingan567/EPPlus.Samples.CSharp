using System;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml.Utils.CompundDocument;

namespace OfficeOpenXml.Encryption;

internal class EncryptedPackageHandler
{
	private readonly byte[] BlockKey_HashInput = new byte[8] { 254, 167, 210, 118, 59, 75, 158, 121 };

	private readonly byte[] BlockKey_HashValue = new byte[8] { 215, 170, 15, 109, 48, 97, 52, 78 };

	private readonly byte[] BlockKey_KeyValue = new byte[8] { 20, 110, 11, 231, 171, 172, 208, 214 };

	private readonly byte[] BlockKey_HmacKey = new byte[8] { 95, 178, 173, 1, 12, 185, 225, 246 };

	private readonly byte[] BlockKey_HmacValue = new byte[8] { 160, 103, 127, 2, 178, 44, 132, 51 };

	internal MemoryStream DecryptPackage(FileInfo fi, ExcelEncryption encryption)
	{
		if (CompoundDocument.IsCompoundDocument(fi))
		{
			CompoundDocument doc = new CompoundDocument(fi);
			return GetStreamFromPackage(doc, encryption);
		}
		throw new InvalidDataException($"File {fi.FullName} is not an encrypted package");
	}

	internal MemoryStream DecryptPackage(MemoryStream stream, ExcelEncryption encryption)
	{
		try
		{
			if (CompoundDocument.IsCompoundDocument(stream))
			{
				CompoundDocument doc = new CompoundDocument(stream);
				return GetStreamFromPackage(doc, encryption);
			}
			throw new InvalidDataException("The stream is not an valid/supported encrypted document.");
		}
		catch
		{
			throw;
		}
	}

	internal MemoryStream EncryptPackage(byte[] package, ExcelEncryption encryption)
	{
		if (encryption.Version == EncryptionVersion.Standard)
		{
			return EncryptPackageBinary(package, encryption);
		}
		if (encryption.Version == EncryptionVersion.Agile)
		{
			return EncryptPackageAgile(package, encryption);
		}
		throw new ArgumentException("Unsupported encryption version.");
	}

	private MemoryStream EncryptPackageAgile(byte[] package, ExcelEncryption encryption)
	{
		string text = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
		text += "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">";
		text += "<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"\"/>";
		text += "<dataIntegrity encryptedHmacKey=\"\" encryptedHmacValue=\"\"/>";
		text += "<keyEncryptors>";
		text += "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">";
		text += "<p:encryptedKey spinCount=\"100000\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"\" encryptedVerifierHashInput=\"\" encryptedVerifierHashValue=\"\" encryptedKeyValue=\"\" />";
		text += "</keyEncryptor></keyEncryptors></encryption>";
		EncryptionInfoAgile encryptionInfoAgile = new EncryptionInfoAgile();
		encryptionInfoAgile.ReadFromXml(text);
		EncryptionInfoAgile.EncryptionKeyEncryptor encryptionKeyEncryptor = encryptionInfoAgile.KeyEncryptors[0];
		RandomNumberGenerator randomNumberGenerator = RandomNumberGenerator.Create();
		byte[] array = new byte[16];
		randomNumberGenerator.GetBytes(array);
		encryptionInfoAgile.KeyData.SaltValue = array;
		randomNumberGenerator.GetBytes(array);
		encryptionKeyEncryptor.SaltValue = array;
		encryptionKeyEncryptor.KeyValue = new byte[encryptionKeyEncryptor.KeyBits / 8];
		randomNumberGenerator.GetBytes(encryptionKeyEncryptor.KeyValue);
		HashAlgorithm hashProvider = GetHashProvider(encryptionInfoAgile.KeyEncryptors[0]);
		byte[] passwordHash = GetPasswordHash(hashProvider, encryptionKeyEncryptor.SaltValue, encryption.Password, encryptionKeyEncryptor.SpinCount, encryptionKeyEncryptor.HashSize);
		byte[] finalHash = GetFinalHash(hashProvider, BlockKey_KeyValue, passwordHash);
		finalHash = FixHashSize(finalHash, encryptionKeyEncryptor.KeyBits / 8, 0);
		byte[] array2 = EncryptDataAgile(package, encryptionInfoAgile, hashProvider);
		byte[] array3 = new byte[64];
		randomNumberGenerator.GetBytes(array3);
		SetHMAC(encryptionInfoAgile, hashProvider, array3, array2);
		encryptionKeyEncryptor.VerifierHashInput = new byte[16];
		randomNumberGenerator.GetBytes(encryptionKeyEncryptor.VerifierHashInput);
		encryptionKeyEncryptor.VerifierHash = hashProvider.ComputeHash(encryptionKeyEncryptor.VerifierHashInput);
		byte[] finalHash2 = GetFinalHash(hashProvider, BlockKey_HashInput, passwordHash);
		byte[] finalHash3 = GetFinalHash(hashProvider, BlockKey_HashValue, passwordHash);
		byte[] finalHash4 = GetFinalHash(hashProvider, BlockKey_KeyValue, passwordHash);
		MemoryStream memoryStream = new MemoryStream();
		EncryptAgileFromKey(encryptionKeyEncryptor, finalHash2, encryptionKeyEncryptor.VerifierHashInput, 0L, encryptionKeyEncryptor.VerifierHashInput.Length, encryptionKeyEncryptor.SaltValue, memoryStream);
		encryptionKeyEncryptor.EncryptedVerifierHashInput = memoryStream.ToArray();
		memoryStream = new MemoryStream();
		EncryptAgileFromKey(encryptionKeyEncryptor, finalHash3, encryptionKeyEncryptor.VerifierHash, 0L, encryptionKeyEncryptor.VerifierHash.Length, encryptionKeyEncryptor.SaltValue, memoryStream);
		encryptionKeyEncryptor.EncryptedVerifierHash = memoryStream.ToArray();
		memoryStream = new MemoryStream();
		EncryptAgileFromKey(encryptionKeyEncryptor, finalHash4, encryptionKeyEncryptor.KeyValue, 0L, encryptionKeyEncryptor.KeyValue.Length, encryptionKeyEncryptor.SaltValue, memoryStream);
		encryptionKeyEncryptor.EncryptedKeyValue = memoryStream.ToArray();
		text = encryptionInfoAgile.Xml.OuterXml;
		byte[] bytes = Encoding.UTF8.GetBytes(text);
		memoryStream = new MemoryStream();
		memoryStream.Write(BitConverter.GetBytes((ushort)4), 0, 2);
		memoryStream.Write(BitConverter.GetBytes((ushort)4), 0, 2);
		memoryStream.Write(BitConverter.GetBytes(64u), 0, 4);
		memoryStream.Write(bytes, 0, bytes.Length);
		CompoundDocument compoundDocument = new CompoundDocument();
		CreateDataSpaces(compoundDocument);
		compoundDocument.Storage.DataStreams.Add("EncryptionInfo", memoryStream.ToArray());
		compoundDocument.Storage.DataStreams.Add("EncryptedPackage", array2);
		memoryStream = new MemoryStream();
		compoundDocument.Save(memoryStream);
		return memoryStream;
	}

	private byte[] EncryptDataAgile(byte[] data, EncryptionInfoAgile encryptionInfo, HashAlgorithm hashProvider)
	{
		EncryptionInfoAgile.EncryptionKeyEncryptor encryptionKeyEncryptor = encryptionInfo.KeyEncryptors[0];
		new RijndaelManaged
		{
			KeySize = encryptionKeyEncryptor.KeyBits,
			Mode = CipherMode.CBC,
			Padding = PaddingMode.Zeros
		};
		int num = 0;
		int num2 = 0;
		MemoryStream memoryStream = new MemoryStream();
		memoryStream.Write(BitConverter.GetBytes((ulong)data.Length), 0, 8);
		while (num < data.Length)
		{
			int num3 = ((data.Length - num > 4096) ? 4096 : (data.Length - num));
			byte[] array = new byte[4 + encryptionInfo.KeyData.SaltSize];
			Array.Copy(encryptionInfo.KeyData.SaltValue, 0, array, 0, encryptionInfo.KeyData.SaltSize);
			Array.Copy(BitConverter.GetBytes(num2), 0, array, encryptionInfo.KeyData.SaltSize, 4);
			byte[] iv = hashProvider.ComputeHash(array);
			EncryptAgileFromKey(encryptionKeyEncryptor, encryptionKeyEncryptor.KeyValue, data, num, num3, iv, memoryStream);
			num += num3;
			num2++;
		}
		memoryStream.Flush();
		return memoryStream.ToArray();
	}

	private void SetHMAC(EncryptionInfoAgile ei, HashAlgorithm hashProvider, byte[] salt, byte[] data)
	{
		byte[] finalHash = GetFinalHash(hashProvider, BlockKey_HmacKey, ei.KeyData.SaltValue);
		MemoryStream memoryStream = new MemoryStream();
		EncryptAgileFromKey(ei.KeyEncryptors[0], ei.KeyEncryptors[0].KeyValue, salt, 0L, salt.Length, finalHash, memoryStream);
		ei.DataIntegrity.EncryptedHmacKey = memoryStream.ToArray();
		byte[] array = GetHmacProvider(ei.KeyEncryptors[0], salt).ComputeHash(data);
		memoryStream = new MemoryStream();
		finalHash = GetFinalHash(hashProvider, BlockKey_HmacValue, ei.KeyData.SaltValue);
		EncryptAgileFromKey(ei.KeyEncryptors[0], ei.KeyEncryptors[0].KeyValue, array, 0L, array.Length, finalHash, memoryStream);
		ei.DataIntegrity.EncryptedHmacValue = memoryStream.ToArray();
	}

	private HMAC GetHmacProvider(EncryptionInfoAgile.EncryptionKeyData ei, byte[] salt)
	{
		return ei.HashAlgorithm switch
		{
			eHashAlogorithm.RIPEMD160 => new HMACRIPEMD160(salt), 
			eHashAlogorithm.MD5 => new HMACMD5(salt), 
			eHashAlogorithm.SHA1 => new HMACSHA1(salt), 
			eHashAlogorithm.SHA256 => new HMACSHA256(salt), 
			eHashAlogorithm.SHA384 => new HMACSHA384(salt), 
			eHashAlogorithm.SHA512 => new HMACSHA512(salt), 
			_ => throw new NotSupportedException($"Hash method {ei.HashAlgorithm} not supported."), 
		};
	}

	private MemoryStream EncryptPackageBinary(byte[] package, ExcelEncryption encryption)
	{
		byte[] key;
		EncryptionInfoBinary encryptionInfoBinary = CreateEncryptionInfo(encryption.Password, (encryption.Algorithm == EncryptionAlgorithm.AES128) ? AlgorithmID.AES128 : ((encryption.Algorithm == EncryptionAlgorithm.AES192) ? AlgorithmID.AES192 : AlgorithmID.AES256), out key);
		CompoundDocument compoundDocument = new CompoundDocument();
		CreateDataSpaces(compoundDocument);
		compoundDocument.Storage.DataStreams.Add("EncryptionInfo", encryptionInfoBinary.WriteBinary());
		byte[] array = EncryptData(key, package, useDataSize: false);
		MemoryStream memoryStream = new MemoryStream();
		memoryStream.Write(BitConverter.GetBytes((ulong)package.Length), 0, 8);
		memoryStream.Write(array, 0, array.Length);
		compoundDocument.Storage.DataStreams.Add("EncryptedPackage", memoryStream.ToArray());
		MemoryStream memoryStream2 = new MemoryStream();
		compoundDocument.Save(memoryStream2);
		return memoryStream2;
	}

	private void CreateDataSpaces(CompoundDocument doc)
	{
		CompoundDocument.StoragePart storagePart = new CompoundDocument.StoragePart();
		doc.Storage.SubStorage.Add("\u0006DataSpaces", storagePart);
		new CompoundDocument.StoragePart();
		storagePart.DataStreams.Add("Version", CreateVersionStream());
		storagePart.DataStreams.Add("DataSpaceMap", CreateDataSpaceMap());
		CompoundDocument.StoragePart storagePart2 = new CompoundDocument.StoragePart();
		storagePart.SubStorage.Add("DataSpaceInfo", storagePart2);
		storagePart2.DataStreams.Add("StrongEncryptionDataSpace", CreateStrongEncryptionDataSpaceStream());
		CompoundDocument.StoragePart storagePart3 = new CompoundDocument.StoragePart();
		storagePart.SubStorage.Add("TransformInfo", storagePart3);
		CompoundDocument.StoragePart storagePart4 = new CompoundDocument.StoragePart();
		storagePart3.SubStorage.Add("StrongEncryptionTransform", storagePart4);
		storagePart4.DataStreams.Add("\u0006Primary", CreateTransformInfoPrimary());
	}

	private byte[] CreateStrongEncryptionDataSpaceStream()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write(8);
		binaryWriter.Write(1);
		string text = "StrongEncryptionTransform";
		binaryWriter.Write(text.Length * 2);
		binaryWriter.Write(Encoding.Unicode.GetBytes(text + "\0"));
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private byte[] CreateVersionStream()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write((short)60);
		binaryWriter.Write((short)0);
		binaryWriter.Write(Encoding.Unicode.GetBytes("Microsoft.Container.DataSpaces"));
		binaryWriter.Write(1);
		binaryWriter.Write(1);
		binaryWriter.Write(1);
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private byte[] CreateDataSpaceMap()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write(8);
		binaryWriter.Write(1);
		string text = "EncryptedPackage";
		string text2 = "StrongEncryptionDataSpace";
		binaryWriter.Write((text.Length + text2.Length) * 2 + 22);
		binaryWriter.Write(1);
		binaryWriter.Write(0);
		binaryWriter.Write(text.Length * 2);
		binaryWriter.Write(Encoding.Unicode.GetBytes(text));
		binaryWriter.Write(text2.Length * 2);
		binaryWriter.Write(Encoding.Unicode.GetBytes(text2 + "\0"));
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private byte[] CreateTransformInfoPrimary()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		string text = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
		string text2 = "Microsoft.Container.EncryptionTransform";
		binaryWriter.Write(text.Length * 2 + 12);
		binaryWriter.Write(1);
		binaryWriter.Write(text.Length * 2);
		binaryWriter.Write(Encoding.Unicode.GetBytes(text));
		binaryWriter.Write(text2.Length * 2);
		binaryWriter.Write(Encoding.Unicode.GetBytes(text2 + "\0"));
		binaryWriter.Write(1);
		binaryWriter.Write(1);
		binaryWriter.Write(1);
		binaryWriter.Write(0);
		binaryWriter.Write(0);
		binaryWriter.Write(0);
		binaryWriter.Write(4);
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private EncryptionInfoBinary CreateEncryptionInfo(string password, AlgorithmID algID, out byte[] key)
	{
		if (algID == AlgorithmID.Flags || algID == AlgorithmID.RC4)
		{
			throw new ArgumentException("algID must be AES128, AES192 or AES256");
		}
		EncryptionInfoBinary encryptionInfoBinary = new EncryptionInfoBinary();
		encryptionInfoBinary.MajorVersion = 4;
		encryptionInfoBinary.MinorVersion = 2;
		encryptionInfoBinary.Flags = Flags.fCryptoAPI | Flags.fAES;
		encryptionInfoBinary.Header = new EncryptionHeader();
		encryptionInfoBinary.Header.AlgID = algID;
		encryptionInfoBinary.Header.AlgIDHash = AlgorithmHashID.SHA1;
		encryptionInfoBinary.Header.Flags = encryptionInfoBinary.Flags;
		encryptionInfoBinary.Header.KeySize = algID switch
		{
			AlgorithmID.AES192 => 192, 
			AlgorithmID.AES128 => 128, 
			_ => 256, 
		};
		encryptionInfoBinary.Header.ProviderType = ProviderType.AES;
		encryptionInfoBinary.Header.CSPName = "Microsoft Enhanced RSA and AES Cryptographic Provider\0";
		encryptionInfoBinary.Header.Reserved1 = 0;
		encryptionInfoBinary.Header.Reserved2 = 0;
		encryptionInfoBinary.Header.SizeExtra = 0;
		encryptionInfoBinary.Verifier = new EncryptionVerifier();
		encryptionInfoBinary.Verifier.Salt = new byte[16];
		RandomNumberGenerator randomNumberGenerator = RandomNumberGenerator.Create();
		randomNumberGenerator.GetBytes(encryptionInfoBinary.Verifier.Salt);
		encryptionInfoBinary.Verifier.SaltSize = 16u;
		key = GetPasswordHashBinary(password, encryptionInfoBinary);
		byte[] array = new byte[16];
		randomNumberGenerator.GetBytes(array);
		encryptionInfoBinary.Verifier.EncryptedVerifier = EncryptData(key, array, useDataSize: true);
		encryptionInfoBinary.Verifier.VerifierHashSize = 32u;
		byte[] data = new SHA1Managed().ComputeHash(array);
		encryptionInfoBinary.Verifier.EncryptedVerifierHash = EncryptData(key, data, useDataSize: false);
		return encryptionInfoBinary;
	}

	private byte[] EncryptData(byte[] key, byte[] data, bool useDataSize)
	{
		ICryptoTransform transform = new RijndaelManaged
		{
			KeySize = key.Length * 8,
			Mode = CipherMode.ECB,
			Padding = PaddingMode.Zeros
		}.CreateEncryptor(key, null);
		MemoryStream memoryStream = new MemoryStream();
		CryptoStream cryptoStream = new CryptoStream(memoryStream, transform, CryptoStreamMode.Write);
		cryptoStream.Write(data, 0, data.Length);
		cryptoStream.FlushFinalBlock();
		if (useDataSize)
		{
			byte[] array = new byte[data.Length];
			memoryStream.Seek(0L, SeekOrigin.Begin);
			memoryStream.Read(array, 0, data.Length);
			return array;
		}
		return memoryStream.ToArray();
	}

	private MemoryStream GetStreamFromPackage(CompoundDocument doc, ExcelEncryption encryption)
	{
		new MemoryStream();
		if (doc.Storage.DataStreams.ContainsKey("EncryptionInfo") || doc.Storage.DataStreams.ContainsKey("EncryptedPackage"))
		{
			EncryptionInfo encryptionInfo = EncryptionInfo.ReadBinary(doc.Storage.DataStreams["EncryptionInfo"]);
			return DecryptDocument(doc.Storage.DataStreams["EncryptedPackage"], encryptionInfo, encryption.Password);
		}
		throw new InvalidDataException("Invalid document. EncryptionInfo or EncryptedPackage stream is missing");
	}

	private MemoryStream DecryptDocument(byte[] data, EncryptionInfo encryptionInfo, string password)
	{
		long size = BitConverter.ToInt64(data, 0);
		byte[] array = new byte[data.Length - 8];
		Array.Copy(data, 8, array, 0, array.Length);
		if (encryptionInfo is EncryptionInfoBinary)
		{
			return DecryptBinary((EncryptionInfoBinary)encryptionInfo, password, size, array);
		}
		return DecryptAgile((EncryptionInfoAgile)encryptionInfo, password, size, array, data);
	}

	private MemoryStream DecryptAgile(EncryptionInfoAgile encryptionInfo, string password, long size, byte[] encryptedData, byte[] data)
	{
		MemoryStream memoryStream = new MemoryStream();
		if (encryptionInfo.KeyData.CipherAlgorithm == eCipherAlgorithm.AES)
		{
			EncryptionInfoAgile.EncryptionKeyEncryptor encryptionKeyEncryptor = encryptionInfo.KeyEncryptors[0];
			HashAlgorithm hashProvider = GetHashProvider(encryptionKeyEncryptor);
			HashAlgorithm hashProvider2 = GetHashProvider(encryptionInfo.KeyData);
			byte[] passwordHash = GetPasswordHash(hashProvider, encryptionKeyEncryptor.SaltValue, password, encryptionKeyEncryptor.SpinCount, encryptionKeyEncryptor.HashSize);
			byte[] finalHash = GetFinalHash(hashProvider, BlockKey_HashInput, passwordHash);
			byte[] finalHash2 = GetFinalHash(hashProvider, BlockKey_HashValue, passwordHash);
			byte[] finalHash3 = GetFinalHash(hashProvider, BlockKey_KeyValue, passwordHash);
			encryptionKeyEncryptor.VerifierHashInput = DecryptAgileFromKey(encryptionKeyEncryptor, finalHash, encryptionKeyEncryptor.EncryptedVerifierHashInput, encryptionKeyEncryptor.SaltSize, encryptionKeyEncryptor.SaltValue);
			encryptionKeyEncryptor.VerifierHash = DecryptAgileFromKey(encryptionKeyEncryptor, finalHash2, encryptionKeyEncryptor.EncryptedVerifierHash, encryptionKeyEncryptor.HashSize, encryptionKeyEncryptor.SaltValue);
			encryptionKeyEncryptor.KeyValue = DecryptAgileFromKey(encryptionKeyEncryptor, finalHash3, encryptionKeyEncryptor.EncryptedKeyValue, encryptionInfo.KeyData.KeyBits / 8, encryptionKeyEncryptor.SaltValue);
			if (IsPasswordValid(hashProvider, encryptionKeyEncryptor))
			{
				byte[] finalHash4 = GetFinalHash(hashProvider2, BlockKey_HmacKey, encryptionInfo.KeyData.SaltValue);
				byte[] salt = DecryptAgileFromKey(encryptionInfo.KeyData, encryptionKeyEncryptor.KeyValue, encryptionInfo.DataIntegrity.EncryptedHmacKey, encryptionInfo.KeyData.HashSize, finalHash4);
				finalHash4 = GetFinalHash(hashProvider2, BlockKey_HmacValue, encryptionInfo.KeyData.SaltValue);
				byte[] array = DecryptAgileFromKey(encryptionInfo.KeyData, encryptionKeyEncryptor.KeyValue, encryptionInfo.DataIntegrity.EncryptedHmacValue, encryptionInfo.KeyData.HashSize, finalHash4);
				byte[] array2 = GetHmacProvider(encryptionInfo.KeyData, salt).ComputeHash(data);
				for (int i = 0; i < array2.Length; i++)
				{
					if (array[i] != array2[i])
					{
						throw new Exception("Dataintegrity key mismatch");
					}
				}
				int num = 0;
				uint num2 = 0u;
				while (num < size)
				{
					int num3 = (int)((size - num > 4096) ? 4096 : (size - num));
					int num4 = ((encryptedData.Length - num > 4096) ? 4096 : (encryptedData.Length - num));
					byte[] array3 = new byte[4 + encryptionInfo.KeyData.SaltSize];
					Array.Copy(encryptionInfo.KeyData.SaltValue, 0, array3, 0, encryptionInfo.KeyData.SaltSize);
					Array.Copy(BitConverter.GetBytes(num2), 0, array3, encryptionInfo.KeyData.SaltSize, 4);
					byte[] iv = hashProvider2.ComputeHash(array3);
					byte[] array4 = new byte[num4];
					Array.Copy(encryptedData, num, array4, 0, num4);
					byte[] array5 = DecryptAgileFromKey(encryptionInfo.KeyData, encryptionKeyEncryptor.KeyValue, array4, num3, iv);
					memoryStream.Write(array5, 0, array5.Length);
					num += num3;
					num2++;
				}
				memoryStream.Flush();
				return memoryStream;
			}
			throw new SecurityException("Invalid password");
		}
		return null;
	}

	private HashAlgorithm GetHashProvider(EncryptionInfoAgile.EncryptionKeyData encr)
	{
		return encr.HashAlgorithm switch
		{
			eHashAlogorithm.MD5 => new MD5CryptoServiceProvider(), 
			eHashAlogorithm.RIPEMD160 => new RIPEMD160Managed(), 
			eHashAlogorithm.SHA1 => new SHA1CryptoServiceProvider(), 
			eHashAlogorithm.SHA256 => new SHA256CryptoServiceProvider(), 
			eHashAlogorithm.SHA384 => new SHA384CryptoServiceProvider(), 
			eHashAlogorithm.SHA512 => new SHA512CryptoServiceProvider(), 
			_ => throw new NotSupportedException($"Hash provider is unsupported. {encr.HashAlgorithm}"), 
		};
	}

	private MemoryStream DecryptBinary(EncryptionInfoBinary encryptionInfo, string password, long size, byte[] encryptedData)
	{
		MemoryStream memoryStream = new MemoryStream();
		if (encryptionInfo.Header.AlgID == AlgorithmID.AES128 || (encryptionInfo.Header.AlgID == AlgorithmID.Flags && (encryptionInfo.Flags & (Flags.fCryptoAPI | Flags.fExternal | Flags.fAES)) == (Flags.fCryptoAPI | Flags.fAES)) || encryptionInfo.Header.AlgID == AlgorithmID.AES192 || encryptionInfo.Header.AlgID == AlgorithmID.AES256)
		{
			RijndaelManaged rijndaelManaged = new RijndaelManaged();
			rijndaelManaged.KeySize = encryptionInfo.Header.KeySize;
			rijndaelManaged.Mode = CipherMode.ECB;
			rijndaelManaged.Padding = PaddingMode.None;
			byte[] passwordHashBinary = GetPasswordHashBinary(password, encryptionInfo);
			if (!IsPasswordValid(passwordHashBinary, encryptionInfo))
			{
				throw new UnauthorizedAccessException("Invalid password");
			}
			ICryptoTransform transform = rijndaelManaged.CreateDecryptor(passwordHashBinary, null);
			CryptoStream cryptoStream = new CryptoStream(new MemoryStream(encryptedData), transform, CryptoStreamMode.Read);
			byte[] buffer = new byte[size];
			cryptoStream.Read(buffer, 0, (int)size);
			memoryStream.Write(buffer, 0, (int)size);
		}
		return memoryStream;
	}

	private bool IsPasswordValid(byte[] key, EncryptionInfoBinary encryptionInfo)
	{
		ICryptoTransform transform = new RijndaelManaged
		{
			KeySize = encryptionInfo.Header.KeySize,
			Mode = CipherMode.ECB,
			Padding = PaddingMode.None
		}.CreateDecryptor(key, null);
		CryptoStream cryptoStream = new CryptoStream(new MemoryStream(encryptionInfo.Verifier.EncryptedVerifier), transform, CryptoStreamMode.Read);
		byte[] buffer = new byte[16];
		cryptoStream.Read(buffer, 0, 16);
		CryptoStream cryptoStream2 = new CryptoStream(new MemoryStream(encryptionInfo.Verifier.EncryptedVerifierHash), transform, CryptoStreamMode.Read);
		byte[] array = new byte[16];
		cryptoStream2.Read(array, 0, 16);
		byte[] array2 = new SHA1Managed().ComputeHash(buffer);
		for (int i = 0; i < 16; i++)
		{
			if (array2[i] != array[i])
			{
				return false;
			}
		}
		return true;
	}

	private bool IsPasswordValid(HashAlgorithm sha, EncryptionInfoAgile.EncryptionKeyEncryptor encr)
	{
		byte[] array = sha.ComputeHash(encr.VerifierHashInput);
		for (int i = 0; i < array.Length; i++)
		{
			if (encr.VerifierHash[i] != array[i])
			{
				return false;
			}
		}
		return true;
	}

	private byte[] DecryptAgileFromKey(EncryptionInfoAgile.EncryptionKeyData encr, byte[] key, byte[] encryptedData, long size, byte[] iv)
	{
		SymmetricAlgorithm encryptionAlgorithm = GetEncryptionAlgorithm(encr);
		encryptionAlgorithm.BlockSize = encr.BlockSize << 3;
		encryptionAlgorithm.KeySize = encr.KeyBits;
		encryptionAlgorithm.Mode = ((encr.CipherChaining == eChainingMode.ChainingModeCBC) ? CipherMode.CBC : CipherMode.CFB);
		encryptionAlgorithm.Padding = PaddingMode.Zeros;
		ICryptoTransform transform = encryptionAlgorithm.CreateDecryptor(FixHashSize(key, encr.KeyBits / 8, 0), FixHashSize(iv, encr.BlockSize, 54));
		CryptoStream cryptoStream = new CryptoStream(new MemoryStream(encryptedData), transform, CryptoStreamMode.Read);
		byte[] array = new byte[size];
		cryptoStream.Read(array, 0, (int)size);
		return array;
	}

	private SymmetricAlgorithm GetEncryptionAlgorithm(EncryptionInfoAgile.EncryptionKeyData encr)
	{
		switch (encr.CipherAlgorithm)
		{
		case eCipherAlgorithm.AES:
			return new RijndaelManaged();
		case eCipherAlgorithm.DES:
			return new DESCryptoServiceProvider();
		case eCipherAlgorithm.TRIPLE_DES:
		case eCipherAlgorithm.TRIPLE_DES_112:
			return new TripleDESCryptoServiceProvider();
		case eCipherAlgorithm.RC2:
			return new RC2CryptoServiceProvider();
		default:
			throw new NotSupportedException($"Unsupported Cipher Algorithm: {encr.CipherAlgorithm.ToString()}");
		}
	}

	private void EncryptAgileFromKey(EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] key, byte[] data, long pos, long size, byte[] iv, MemoryStream ms)
	{
		SymmetricAlgorithm encryptionAlgorithm = GetEncryptionAlgorithm(encr);
		encryptionAlgorithm.BlockSize = encr.BlockSize << 3;
		encryptionAlgorithm.KeySize = encr.KeyBits;
		encryptionAlgorithm.Mode = ((encr.CipherChaining == eChainingMode.ChainingModeCBC) ? CipherMode.CBC : CipherMode.CFB);
		encryptionAlgorithm.Padding = PaddingMode.Zeros;
		ICryptoTransform transform = encryptionAlgorithm.CreateEncryptor(FixHashSize(key, encr.KeyBits / 8, 0), FixHashSize(iv, 16, 54));
		CryptoStream cryptoStream = new CryptoStream(ms, transform, CryptoStreamMode.Write);
		if (size % encr.BlockSize != 0L)
		{
			_ = encr.BlockSize;
			_ = size % encr.BlockSize;
		}
		byte[] array = new byte[size];
		Array.Copy(data, (int)pos, array, 0, (int)size);
		cryptoStream.Write(array, 0, (int)size);
		while (size % encr.BlockSize != 0L)
		{
			cryptoStream.WriteByte(0);
			size++;
		}
	}

	private byte[] GetPasswordHashBinary(string password, EncryptionInfoBinary encryptionInfo)
	{
		byte[] array = null;
		byte[] array2 = new byte[24];
		try
		{
			if (encryptionInfo.Header.AlgIDHash == AlgorithmHashID.SHA1 || (encryptionInfo.Header.AlgIDHash == AlgorithmHashID.App && (encryptionInfo.Flags & Flags.fExternal) == 0))
			{
				HashAlgorithm hashAlgorithm = new SHA1CryptoServiceProvider();
				array = GetPasswordHash(hashAlgorithm, encryptionInfo.Verifier.Salt, password, 50000, 20);
				Array.Copy(array, array2, array.Length);
				Array.Copy(BitConverter.GetBytes(0), 0, array2, array.Length, 4);
				array = hashAlgorithm.ComputeHash(array2);
				byte[] array3 = new byte[64];
				int num = encryptionInfo.Header.KeySize / 8;
				for (int i = 0; i < array3.Length; i++)
				{
					array3[i] = (byte)((i < array.Length) ? (0x36u ^ array[i]) : 54u);
				}
				byte[] array4 = hashAlgorithm.ComputeHash(array3);
				if ((int)encryptionInfo.Verifier.VerifierHashSize > num)
				{
					return FixHashSize(array4, num, 0);
				}
				for (int j = 0; j < array3.Length; j++)
				{
					array3[j] = (byte)((j < array.Length) ? (0x5Cu ^ array[j]) : 92u);
				}
				byte[] array5 = hashAlgorithm.ComputeHash(array3);
				byte[] array6 = new byte[array4.Length + array5.Length];
				Array.Copy(array4, 0, array6, 0, array4.Length);
				Array.Copy(array5, 0, array6, array4.Length, array5.Length);
				return FixHashSize(array6, num, 0);
			}
			if (encryptionInfo.Header.KeySize > 0 && encryptionInfo.Header.KeySize < 80)
			{
				throw new NotSupportedException("RC4 Hash provider is not supported. Must be SHA1(AlgIDHash == 0x8004)");
			}
			throw new NotSupportedException("Hash provider is invalid. Must be SHA1(AlgIDHash == 0x8004)");
		}
		catch (Exception innerException)
		{
			throw new Exception("An error occured when the encryptionkey was created", innerException);
		}
	}

	private byte[] GetPasswordHashAgile(string password, EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] blockKey)
	{
		try
		{
			HashAlgorithm hashProvider = GetHashProvider(encr);
			byte[] passwordHash = GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);
			byte[] finalHash = GetFinalHash(hashProvider, blockKey, passwordHash);
			return FixHashSize(finalHash, encr.KeyBits / 8, 0);
		}
		catch (Exception innerException)
		{
			throw new Exception("An error occured when the encryptionkey was created", innerException);
		}
	}

	private byte[] GetFinalHash(HashAlgorithm hashProvider, byte[] blockKey, byte[] hash)
	{
		byte[] array = new byte[hash.Length + blockKey.Length];
		Array.Copy(hash, array, hash.Length);
		Array.Copy(blockKey, 0, array, hash.Length, blockKey.Length);
		return hashProvider.ComputeHash(array);
	}

	private byte[] GetPasswordHash(HashAlgorithm hashProvider, byte[] salt, string password, int spinCount, int hashSize)
	{
		byte[] array = null;
		byte[] array2 = new byte[4 + hashSize];
		array = hashProvider.ComputeHash(CombinePassword(salt, password));
		for (int i = 0; i < spinCount; i++)
		{
			Array.Copy(BitConverter.GetBytes(i), array2, 4);
			Array.Copy(array, 0, array2, 4, array.Length);
			array = hashProvider.ComputeHash(array2);
		}
		return array;
	}

	private byte[] FixHashSize(byte[] hash, int size, byte fill = 0)
	{
		if (hash.Length == size)
		{
			return hash;
		}
		if (hash.Length < size)
		{
			byte[] array = new byte[size];
			Array.Copy(hash, array, hash.Length);
			for (int i = hash.Length; i < size; i++)
			{
				array[i] = fill;
			}
			return array;
		}
		byte[] array2 = new byte[size];
		Array.Copy(hash, array2, size);
		return array2;
	}

	private byte[] CombinePassword(byte[] salt, string password)
	{
		if (password == "")
		{
			password = "VelvetSweatshop";
		}
		byte[] bytes = Encoding.Unicode.GetBytes(password);
		byte[] array = new byte[salt.Length + bytes.Length];
		Array.Copy(salt, array, salt.Length);
		Array.Copy(bytes, 0, array, salt.Length, bytes.Length);
		return array;
	}

	internal static ushort CalculatePasswordHash(string Password)
	{
		ushort num = 0;
		for (int num2 = Password.Length - 1; num2 >= 0; num2--)
		{
			num ^= Password[num2];
			num = (ushort)((ushort)((uint)(num >> 14) & 1u) | (ushort)((uint)(num << 1) & 0x7FFFu));
		}
		num = (ushort)(num ^ 0xCE4Bu);
		return (ushort)(num ^ (ushort)Password.Length);
	}
}
