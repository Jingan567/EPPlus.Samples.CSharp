using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation;

public abstract class ExcelDataValidation : XmlHelper, IExcelDataValidation
{
	private const string _itemElementNodeName = "d:dataValidation";

	private readonly string _errorStylePath = "@errorStyle";

	private readonly string _errorTitlePath = "@errorTitle";

	private readonly string _errorPath = "@error";

	private readonly string _promptTitlePath = "@promptTitle";

	private readonly string _promptPath = "@prompt";

	private readonly string _operatorPath = "@operator";

	private readonly string _showErrorMessagePath = "@showErrorMessage";

	private readonly string _showInputMessagePath = "@showInputMessage";

	private readonly string _typeMessagePath = "@type";

	private readonly string _sqrefPath = "@sqref";

	private readonly string _allowBlankPath = "@allowBlank";

	protected readonly string _formula1Path = "d:formula1";

	protected readonly string _formula2Path = "d:formula2";

	public bool AllowsOperator => ValidationType.AllowOperator;

	public ExcelAddress Address
	{
		get
		{
			return new ExcelAddress(GetXmlNodeString(_sqrefPath));
		}
		private set
		{
			SetAddress(value.Address);
		}
	}

	public ExcelDataValidationType ValidationType
	{
		get
		{
			return ExcelDataValidationType.GetBySchemaName(GetXmlNodeString(_typeMessagePath));
		}
		private set
		{
			SetXmlNodeString(_typeMessagePath, value.SchemaName, removeIfBlank: true);
		}
	}

	public ExcelDataValidationOperator Operator
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(_operatorPath);
			if (!string.IsNullOrEmpty(xmlNodeString))
			{
				return (ExcelDataValidationOperator)Enum.Parse(typeof(ExcelDataValidationOperator), xmlNodeString);
			}
			return ExcelDataValidationOperator.any;
		}
		set
		{
			if (!ValidationType.AllowOperator)
			{
				throw new InvalidOperationException("The current validation type does not allow operator to be set");
			}
			SetXmlNodeString(_operatorPath, value.ToString());
		}
	}

	public ExcelDataValidationWarningStyle ErrorStyle
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(_errorStylePath);
			if (!string.IsNullOrEmpty(xmlNodeString))
			{
				return (ExcelDataValidationWarningStyle)Enum.Parse(typeof(ExcelDataValidationWarningStyle), xmlNodeString);
			}
			return ExcelDataValidationWarningStyle.undefined;
		}
		set
		{
			if (value == ExcelDataValidationWarningStyle.undefined)
			{
				DeleteNode(_errorStylePath);
			}
			SetXmlNodeString(_errorStylePath, value.ToString());
		}
	}

	public bool? AllowBlank
	{
		get
		{
			return GetXmlNodeBoolNullable(_allowBlankPath);
		}
		set
		{
			SetNullableBoolValue(_allowBlankPath, value);
		}
	}

	public bool? ShowInputMessage
	{
		get
		{
			return GetXmlNodeBoolNullable(_showInputMessagePath);
		}
		set
		{
			SetNullableBoolValue(_showInputMessagePath, value);
		}
	}

	public bool? ShowErrorMessage
	{
		get
		{
			return GetXmlNodeBoolNullable(_showErrorMessagePath);
		}
		set
		{
			SetNullableBoolValue(_showErrorMessagePath, value);
		}
	}

	public string ErrorTitle
	{
		get
		{
			return GetXmlNodeString(_errorTitlePath);
		}
		set
		{
			SetXmlNodeString(_errorTitlePath, value);
		}
	}

	public string Error
	{
		get
		{
			return GetXmlNodeString(_errorPath);
		}
		set
		{
			SetXmlNodeString(_errorPath, value);
		}
	}

	public string PromptTitle
	{
		get
		{
			return GetXmlNodeString(_promptTitlePath);
		}
		set
		{
			SetXmlNodeString(_promptTitlePath, value);
		}
	}

	public string Prompt
	{
		get
		{
			return GetXmlNodeString(_promptPath);
		}
		set
		{
			SetXmlNodeString(_promptPath, value);
		}
	}

	protected string Formula1Internal => GetXmlNodeString(_formula1Path);

	protected string Formula2Internal => GetXmlNodeString(_formula2Path);

	internal ExcelDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: this(worksheet, address, validationType, null)
	{
	}

	internal ExcelDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: this(worksheet, address, validationType, itemElementNode, null)
	{
	}

	internal ExcelDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base((namespaceManager != null) ? namespaceManager : worksheet.NameSpaceManager)
	{
		Require.Argument(address).IsNotNullOrEmpty("address");
		address = CheckAndFixRangeAddress(address);
		if (itemElementNode == null)
		{
			base.TopNode = worksheet.WorksheetXml.SelectSingleNode("//d:dataValidations", worksheet.NameSpaceManager);
			string namespaceURI = base.NameSpaceManager.LookupNamespace("d");
			itemElementNode = base.TopNode.OwnerDocument.CreateElement("d:dataValidation".Split(':')[1], namespaceURI);
			base.TopNode.AppendChild(itemElementNode);
		}
		base.TopNode = itemElementNode;
		ValidationType = validationType;
		Address = new ExcelAddress(address);
		Init();
	}

	private void Init()
	{
		base.SchemaNodeOrder = new string[13]
		{
			"type", "errorStyle", "operator", "allowBlank", "showInputMessage", "showErrorMessage", "errorTitle", "error", "promptTitle", "prompt",
			"sqref", "formula1", "formula2"
		};
	}

	private string CheckAndFixRangeAddress(string address)
	{
		if (address.Contains(','))
		{
			throw new FormatException("Multiple addresses may not be commaseparated, use space instead");
		}
		address = ConvertUtil._invariantTextInfo.ToUpper(address);
		if (Regex.IsMatch(address, "[A-Z]+:[A-Z]+"))
		{
			address = AddressUtility.ParseEntireColumnSelections(address);
		}
		return address;
	}

	private void SetNullableBoolValue(string path, bool? val)
	{
		if (val.HasValue)
		{
			SetXmlNodeBool(path, val.Value);
		}
		else
		{
			DeleteNode(path);
		}
	}

	public virtual void Validate()
	{
		string address = Address.Address;
		if (string.IsNullOrEmpty(Formula1Internal))
		{
			throw new InvalidOperationException("Validation of " + address + " failed: Formula1 cannot be empty");
		}
	}

	protected void SetValue<T>(T? val, string path) where T : struct
	{
		if (!val.HasValue)
		{
			DeleteNode(path);
		}
		string value = val.Value.ToString().Replace(',', '.');
		SetXmlNodeString(path, value);
	}

	internal void SetAddress(string address)
	{
		string value = AddressUtility.ParseEntireColumnSelections(address);
		SetXmlNodeString(_sqrefPath, value);
	}
}
