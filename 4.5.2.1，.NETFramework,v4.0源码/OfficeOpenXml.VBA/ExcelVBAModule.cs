using System;
using System.Linq;

namespace OfficeOpenXml.VBA;

public class ExcelVBAModule
{
	private string _name = "";

	private ModuleNameChange _nameChangeCallback;

	private string _code = "";

	public string Name
	{
		get
		{
			return _name;
		}
		set
		{
			if (value.Any((char c) => c > 'ÿ'))
			{
				throw new InvalidOperationException("Vba module names can't contain unicode characters");
			}
			if (value != _name)
			{
				_name = value;
				streamName = value;
				if (_nameChangeCallback != null)
				{
					_nameChangeCallback(value);
				}
			}
		}
	}

	public string Description { get; set; }

	public string Code
	{
		get
		{
			return _code;
		}
		set
		{
			if (value.StartsWith("Attribute", StringComparison.OrdinalIgnoreCase) || value.StartsWith("VERSION", StringComparison.OrdinalIgnoreCase))
			{
				throw new InvalidOperationException("Code can't start with an Attribute or VERSION keyword. Attributes can be accessed through the Attributes collection.");
			}
			_code = value;
		}
	}

	public int HelpContext { get; set; }

	public ExcelVbaModuleAttributesCollection Attributes { get; internal set; }

	public eModuleType Type { get; internal set; }

	public bool ReadOnly { get; set; }

	public bool Private { get; set; }

	internal string streamName { get; set; }

	internal ushort Cookie { get; set; }

	internal uint ModuleOffset { get; set; }

	internal string ClassID { get; set; }

	internal ExcelVBAModule()
	{
		Attributes = new ExcelVbaModuleAttributesCollection();
	}

	internal ExcelVBAModule(ModuleNameChange nameChangeCallback)
		: this()
	{
		_nameChangeCallback = nameChangeCallback;
	}

	public override string ToString()
	{
		return Name;
	}
}
