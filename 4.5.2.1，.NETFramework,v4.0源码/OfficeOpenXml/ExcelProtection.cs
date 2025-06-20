using System.Xml;
using OfficeOpenXml.Encryption;

namespace OfficeOpenXml;

public class ExcelProtection : XmlHelper
{
	private const string workbookPasswordPath = "d:workbookProtection/@workbookPassword";

	private const string lockStructurePath = "d:workbookProtection/@lockStructure";

	private const string lockWindowsPath = "d:workbookProtection/@lockWindows";

	private const string lockRevisionPath = "d:workbookProtection/@lockRevision";

	public bool LockStructure
	{
		get
		{
			return GetXmlNodeBool("d:workbookProtection/@lockStructure", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:workbookProtection/@lockStructure", value, removeIf: false);
		}
	}

	public bool LockWindows
	{
		get
		{
			return GetXmlNodeBool("d:workbookProtection/@lockWindows", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:workbookProtection/@lockWindows", value, removeIf: false);
		}
	}

	public bool LockRevision
	{
		get
		{
			return GetXmlNodeBool("d:workbookProtection/@lockRevision", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:workbookProtection/@lockRevision", value, removeIf: false);
		}
	}

	internal ExcelProtection(XmlNamespaceManager ns, XmlNode topNode, ExcelWorkbook wb)
		: base(ns, topNode)
	{
		base.SchemaNodeOrder = wb.SchemaNodeOrder;
	}

	public void SetPassword(string Password)
	{
		if (string.IsNullOrEmpty(Password))
		{
			DeleteNode("d:workbookProtection/@workbookPassword");
		}
		else
		{
			SetXmlNodeString("d:workbookProtection/@workbookPassword", ((int)EncryptedPackageHandler.CalculatePasswordHash(Password)).ToString("x"));
		}
	}
}
