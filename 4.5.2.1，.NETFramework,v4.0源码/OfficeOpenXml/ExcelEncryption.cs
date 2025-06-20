namespace OfficeOpenXml;

public class ExcelEncryption
{
	private bool _isEncrypted;

	private string _password;

	private EncryptionVersion _version = EncryptionVersion.Agile;

	public bool IsEncrypted
	{
		get
		{
			return _isEncrypted;
		}
		set
		{
			_isEncrypted = value;
			if (_isEncrypted)
			{
				if (_password == null)
				{
					_password = "";
				}
			}
			else
			{
				_password = null;
			}
		}
	}

	public string Password
	{
		get
		{
			return _password;
		}
		set
		{
			_password = value;
			_isEncrypted = value != null;
		}
	}

	public EncryptionAlgorithm Algorithm { get; set; }

	public EncryptionVersion Version
	{
		get
		{
			return _version;
		}
		set
		{
			if (value != Version)
			{
				if (value == EncryptionVersion.Agile)
				{
					Algorithm = EncryptionAlgorithm.AES256;
				}
				else
				{
					Algorithm = EncryptionAlgorithm.AES128;
				}
				_version = value;
			}
		}
	}

	internal ExcelEncryption()
	{
		Algorithm = EncryptionAlgorithm.AES256;
	}

	internal ExcelEncryption(EncryptionAlgorithm encryptionAlgorithm)
	{
		Algorithm = encryptionAlgorithm;
	}
}
