using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public class ExcelProtectedRange : XmlHelper
{
	private ExcelAddress _address;

	public string Name
	{
		get
		{
			return GetXmlNodeString("@name");
		}
		set
		{
			SetXmlNodeString("@name", value);
		}
	}

	public ExcelAddress Address
	{
		get
		{
			if (_address == null)
			{
				_address = new ExcelAddress(GetXmlNodeString("@sqref"));
			}
			return _address;
		}
		set
		{
			SetXmlNodeString("@sqref", SqRefUtility.ToSqRefAddress(value.Address));
			_address = value;
		}
	}

	public string SecurityDescriptor
	{
		get
		{
			return GetXmlNodeString("@securityDescriptor");
		}
		set
		{
			SetXmlNodeString("@securityDescriptor", value);
		}
	}

	internal int SpinCount
	{
		get
		{
			return GetXmlNodeInt("@spinCount");
		}
		set
		{
			SetXmlNodeString("@spinCount", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal string Salt
	{
		get
		{
			return GetXmlNodeString("@saltValue");
		}
		set
		{
			SetXmlNodeString("@saltValue", value);
		}
	}

	internal string Hash
	{
		get
		{
			return GetXmlNodeString("@hashValue");
		}
		set
		{
			SetXmlNodeString("@hashValue", value);
		}
	}

	internal eProtectedRangeAlgorithm Algorithm
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@algorithmName");
			return (eProtectedRangeAlgorithm)Enum.Parse(typeof(eProtectedRangeAlgorithm), xmlNodeString.Replace("-", ""));
		}
		set
		{
			string text = value.ToString();
			if (text.StartsWith("SHA"))
			{
				text = text.Insert(3, "-");
			}
			else if (text.StartsWith("RIPEMD"))
			{
				text = text.Insert(6, "-");
			}
			SetXmlNodeString("@algorithmName", text);
		}
	}

	internal ExcelProtectedRange(string name, ExcelAddress address, XmlNamespaceManager ns, XmlNode topNode)
		: base(ns, topNode)
	{
		Name = name;
		Address = address;
	}

	public void SetPassword(string password)
	{
		byte[] bytes = Encoding.Unicode.GetBytes(password);
		RandomNumberGenerator randomNumberGenerator = RandomNumberGenerator.Create();
		byte[] array = new byte[16];
		randomNumberGenerator.GetBytes(array);
		Algorithm = eProtectedRangeAlgorithm.SHA512;
		SpinCount = ((SpinCount < 100000) ? 100000 : SpinCount);
		SHA512CryptoServiceProvider sHA512CryptoServiceProvider = new SHA512CryptoServiceProvider();
		byte[] array2 = new byte[bytes.Length + array.Length];
		Array.Copy(array, array2, array.Length);
		Array.Copy(bytes, 0, array2, 16, bytes.Length);
		byte[] array3 = sHA512CryptoServiceProvider.ComputeHash(array2);
		for (int i = 0; i < SpinCount; i++)
		{
			array2 = new byte[array3.Length + 4];
			Array.Copy(array3, array2, array3.Length);
			Array.Copy(BitConverter.GetBytes(i), 0, array2, array3.Length, 4);
			array3 = sHA512CryptoServiceProvider.ComputeHash(array2);
		}
		Salt = Convert.ToBase64String(array);
		Hash = Convert.ToBase64String(array3);
	}
}
