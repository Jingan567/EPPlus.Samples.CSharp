using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Encryption;

internal class EncryptionInfoAgile : EncryptionInfo
{
	internal class EncryptionKeyData : XmlHelper
	{
		internal byte[] SaltValue
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@saltValue");
				if (!string.IsNullOrEmpty(xmlNodeString))
				{
					return Convert.FromBase64String(xmlNodeString);
				}
				return null;
			}
			set
			{
				SetXmlNodeString("@saltValue", Convert.ToBase64String(value));
			}
		}

		internal eHashAlogorithm HashAlgorithm
		{
			get
			{
				return GetHashAlgorithm(GetXmlNodeString("@hashAlgorithm"));
			}
			set
			{
				SetXmlNodeString("@hashAlgorithm", GetHashAlgorithmString(value));
			}
		}

		internal eChainingMode CipherChaining
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@cipherChaining");
				try
				{
					return (eChainingMode)Enum.Parse(typeof(eChainingMode), xmlNodeString);
				}
				catch
				{
					throw new InvalidDataException("Invalid chaining mode");
				}
			}
			set
			{
				SetXmlNodeString("@cipherChaining", value.ToString());
			}
		}

		internal eCipherAlgorithm CipherAlgorithm
		{
			get
			{
				return GetCipherAlgorithm(GetXmlNodeString("@cipherAlgorithm"));
			}
			set
			{
				SetXmlNodeString("@cipherAlgorithm", GetCipherAlgorithmString(value));
			}
		}

		internal int HashSize
		{
			get
			{
				return GetXmlNodeInt("@hashSize");
			}
			set
			{
				SetXmlNodeString("@hashSize", value.ToString());
			}
		}

		internal int KeyBits
		{
			get
			{
				return GetXmlNodeInt("@keyBits");
			}
			set
			{
				SetXmlNodeString("@keyBits", value.ToString());
			}
		}

		internal int BlockSize
		{
			get
			{
				return GetXmlNodeInt("@blockSize");
			}
			set
			{
				SetXmlNodeString("@blockSize", value.ToString());
			}
		}

		internal int SaltSize
		{
			get
			{
				return GetXmlNodeInt("@saltSize");
			}
			set
			{
				SetXmlNodeString("@saltSize", value.ToString());
			}
		}

		public EncryptionKeyData(XmlNamespaceManager nsm, XmlNode topNode)
			: base(nsm, topNode)
		{
		}

		private eHashAlogorithm GetHashAlgorithm(string v)
		{
			switch (v)
			{
			case "RIPEMD-128":
				return eHashAlogorithm.RIPEMD128;
			case "RIPEMD-160":
				return eHashAlogorithm.RIPEMD160;
			case "SHA-1":
				return eHashAlogorithm.SHA1;
			default:
				try
				{
					return (eHashAlogorithm)Enum.Parse(typeof(eHashAlogorithm), v);
				}
				catch
				{
					throw new InvalidDataException("Invalid Hash algorithm");
				}
			}
		}

		private string GetHashAlgorithmString(eHashAlogorithm value)
		{
			return value switch
			{
				eHashAlogorithm.RIPEMD128 => "RIPEMD-128", 
				eHashAlogorithm.RIPEMD160 => "RIPEMD-160", 
				eHashAlogorithm.SHA1 => "SHA-1", 
				_ => value.ToString(), 
			};
		}

		private eCipherAlgorithm GetCipherAlgorithm(string v)
		{
			if (!(v == "3DES"))
			{
				if (v == "3DES_112")
				{
					return eCipherAlgorithm.TRIPLE_DES_112;
				}
				try
				{
					return (eCipherAlgorithm)Enum.Parse(typeof(eCipherAlgorithm), v);
				}
				catch
				{
					throw new InvalidDataException("Invalid Hash algorithm");
				}
			}
			return eCipherAlgorithm.TRIPLE_DES;
		}

		private string GetCipherAlgorithmString(eCipherAlgorithm alg)
		{
			return alg switch
			{
				eCipherAlgorithm.TRIPLE_DES => "3DES", 
				eCipherAlgorithm.TRIPLE_DES_112 => "3DES_112", 
				_ => alg.ToString(), 
			};
		}
	}

	internal class EncryptionDataIntegrity : XmlHelper
	{
		internal byte[] EncryptedHmacValue
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@encryptedHmacValue");
				if (!string.IsNullOrEmpty(xmlNodeString))
				{
					return Convert.FromBase64String(xmlNodeString);
				}
				return null;
			}
			set
			{
				SetXmlNodeString("@encryptedHmacValue", Convert.ToBase64String(value));
			}
		}

		internal byte[] EncryptedHmacKey
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@encryptedHmacKey");
				if (!string.IsNullOrEmpty(xmlNodeString))
				{
					return Convert.FromBase64String(xmlNodeString);
				}
				return null;
			}
			set
			{
				SetXmlNodeString("@encryptedHmacKey", Convert.ToBase64String(value));
			}
		}

		public EncryptionDataIntegrity(XmlNamespaceManager nsm, XmlNode topNode)
			: base(nsm, topNode)
		{
		}
	}

	internal class EncryptionKeyEncryptor : EncryptionKeyData
	{
		internal byte[] EncryptedKeyValue
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@encryptedKeyValue");
				if (!string.IsNullOrEmpty(xmlNodeString))
				{
					return Convert.FromBase64String(xmlNodeString);
				}
				return null;
			}
			set
			{
				SetXmlNodeString("@encryptedKeyValue", Convert.ToBase64String(value));
			}
		}

		internal byte[] EncryptedVerifierHash
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@encryptedVerifierHashValue");
				if (!string.IsNullOrEmpty(xmlNodeString))
				{
					return Convert.FromBase64String(xmlNodeString);
				}
				return null;
			}
			set
			{
				SetXmlNodeString("@encryptedVerifierHashValue", Convert.ToBase64String(value));
			}
		}

		internal byte[] EncryptedVerifierHashInput
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@encryptedVerifierHashInput");
				if (!string.IsNullOrEmpty(xmlNodeString))
				{
					return Convert.FromBase64String(xmlNodeString);
				}
				return null;
			}
			set
			{
				SetXmlNodeString("@encryptedVerifierHashInput", Convert.ToBase64String(value));
			}
		}

		internal byte[] VerifierHashInput { get; set; }

		internal byte[] VerifierHash { get; set; }

		internal byte[] KeyValue { get; set; }

		internal int SpinCount
		{
			get
			{
				return GetXmlNodeInt("@spinCount");
			}
			set
			{
				SetXmlNodeString("@spinCount", value.ToString());
			}
		}

		public EncryptionKeyEncryptor(XmlNamespaceManager nsm, XmlNode topNode)
			: base(nsm, topNode)
		{
		}
	}

	private XmlNamespaceManager _nsm;

	internal EncryptionDataIntegrity DataIntegrity { get; set; }

	internal EncryptionKeyData KeyData { get; set; }

	internal List<EncryptionKeyEncryptor> KeyEncryptors { get; private set; }

	internal XmlDocument Xml { get; set; }

	public EncryptionInfoAgile()
	{
		NameTable nameTable = new NameTable();
		_nsm = new XmlNamespaceManager(nameTable);
		_nsm.AddNamespace("d", "http://schemas.microsoft.com/office/2006/encryption");
		_nsm.AddNamespace("c", "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");
		_nsm.AddNamespace("p", "http://schemas.microsoft.com/office/2006/keyEncryptor/password");
	}

	internal override void Read(byte[] data)
	{
		byte[] array = new byte[data.Length - 8];
		Array.Copy(data, 8, array, 0, data.Length - 8);
		string @string = Encoding.UTF8.GetString(array);
		ReadFromXml(@string);
	}

	internal void ReadFromXml(string xml)
	{
		Xml = new XmlDocument();
		XmlHelper.LoadXmlSafe(Xml, xml, Encoding.UTF8);
		XmlNode topNode = Xml.SelectSingleNode("/d:encryption/d:keyData", _nsm);
		KeyData = new EncryptionKeyData(_nsm, topNode);
		topNode = Xml.SelectSingleNode("/d:encryption/d:dataIntegrity", _nsm);
		DataIntegrity = new EncryptionDataIntegrity(_nsm, topNode);
		KeyEncryptors = new List<EncryptionKeyEncryptor>();
		XmlNodeList xmlNodeList = Xml.SelectNodes("/d:encryption/d:keyEncryptors/d:keyEncryptor/p:encryptedKey", _nsm);
		if (xmlNodeList == null)
		{
			return;
		}
		foreach (XmlNode item in xmlNodeList)
		{
			KeyEncryptors.Add(new EncryptionKeyEncryptor(_nsm, item));
		}
	}
}
