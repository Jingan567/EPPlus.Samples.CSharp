using System.Xml;
using OfficeOpenXml.Encryption;

namespace OfficeOpenXml;

public sealed class ExcelSheetProtection : XmlHelper
{
	private const string _isProtectedPath = "d:sheetProtection/@sheet";

	private const string _allowSelectLockedCellsPath = "d:sheetProtection/@selectLockedCells";

	private const string _allowSelectUnlockedCellsPath = "d:sheetProtection/@selectUnlockedCells";

	private const string _allowObjectPath = "d:sheetProtection/@objects";

	private const string _allowScenariosPath = "d:sheetProtection/@scenarios";

	private const string _allowFormatCellsPath = "d:sheetProtection/@formatCells";

	private const string _allowFormatColumnsPath = "d:sheetProtection/@formatColumns";

	private const string _allowFormatRowsPath = "d:sheetProtection/@formatRows";

	private const string _allowInsertColumnsPath = "d:sheetProtection/@insertColumns";

	private const string _allowInsertRowsPath = "d:sheetProtection/@insertRows";

	private const string _allowInsertHyperlinksPath = "d:sheetProtection/@insertHyperlinks";

	private const string _allowDeleteColumns = "d:sheetProtection/@deleteColumns";

	private const string _allowDeleteRowsPath = "d:sheetProtection/@deleteRows";

	private const string _allowSortPath = "d:sheetProtection/@sort";

	private const string _allowAutoFilterPath = "d:sheetProtection/@autoFilter";

	private const string _allowPivotTablesPath = "d:sheetProtection/@pivotTables";

	private const string _passwordPath = "d:sheetProtection/@password";

	public bool IsProtected
	{
		get
		{
			return GetXmlNodeBool("d:sheetProtection/@sheet", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@sheet", value, removeIf: false);
			if (value)
			{
				AllowEditObject = true;
				AllowEditScenarios = true;
			}
			else
			{
				DeleteAllNode("d:sheetProtection/@sheet");
			}
		}
	}

	public bool AllowSelectLockedCells
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@selectLockedCells", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@selectLockedCells", !value, removeIf: false);
		}
	}

	public bool AllowSelectUnlockedCells
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@selectUnlockedCells", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@selectUnlockedCells", !value, removeIf: false);
		}
	}

	public bool AllowEditObject
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@objects", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@objects", !value, removeIf: false);
		}
	}

	public bool AllowEditScenarios
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@scenarios", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@scenarios", !value, removeIf: false);
		}
	}

	public bool AllowFormatCells
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@formatCells", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@formatCells", !value, removeIf: true);
		}
	}

	public bool AllowFormatColumns
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@formatColumns", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@formatColumns", !value, removeIf: true);
		}
	}

	public bool AllowFormatRows
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@formatRows", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@formatRows", !value, removeIf: true);
		}
	}

	public bool AllowInsertColumns
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@insertColumns", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@insertColumns", !value, removeIf: true);
		}
	}

	public bool AllowInsertRows
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@insertRows", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@insertRows", !value, removeIf: true);
		}
	}

	public bool AllowInsertHyperlinks
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@insertHyperlinks", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@insertHyperlinks", !value, removeIf: true);
		}
	}

	public bool AllowDeleteColumns
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@deleteColumns", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@deleteColumns", !value, removeIf: true);
		}
	}

	public bool AllowDeleteRows
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@deleteRows", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@deleteRows", !value, removeIf: true);
		}
	}

	public bool AllowSort
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@sort", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@sort", !value, removeIf: true);
		}
	}

	public bool AllowAutoFilter
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@autoFilter", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@autoFilter", !value, removeIf: true);
		}
	}

	public bool AllowPivotTables
	{
		get
		{
			return !GetXmlNodeBool("d:sheetProtection/@pivotTables", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:sheetProtection/@pivotTables", !value, removeIf: true);
		}
	}

	internal ExcelSheetProtection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
		: base(nsm, topNode)
	{
		base.SchemaNodeOrder = ws.SchemaNodeOrder;
	}

	public void SetPassword(string Password)
	{
		if (!IsProtected)
		{
			IsProtected = true;
		}
		Password = Password.Trim();
		if (Password == "")
		{
			XmlNode xmlNode = base.TopNode.SelectSingleNode("d:sheetProtection/@password", base.NameSpaceManager);
			if (xmlNode != null)
			{
				(xmlNode as XmlAttribute).OwnerElement.Attributes.Remove(xmlNode as XmlAttribute);
			}
		}
		else
		{
			SetXmlNodeString("d:sheetProtection/@password", ((int)EncryptedPackageHandler.CalculatePasswordHash(Password)).ToString("x"));
		}
	}
}
