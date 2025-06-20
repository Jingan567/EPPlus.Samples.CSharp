using System.Xml;

namespace OfficeOpenXml.Sparkline;

public class ExcelSparkline : XmlHelper
{
	private const string _fPath = "xm:f";

	private const string _sqrefPath = "xm:sqref";

	public ExcelAddressBase RangeAddress
	{
		get
		{
			return new ExcelAddressBase(GetXmlNodeString("xm:f"));
		}
		internal set
		{
			SetXmlNodeString("xm:f", value.FullAddress);
		}
	}

	public ExcelCellAddress Cell
	{
		get
		{
			return new ExcelCellAddress(GetXmlNodeString("xm:sqref"));
		}
		internal set
		{
			SetXmlNodeString("xm:sqref", value.Address);
		}
	}

	internal ExcelSparkline(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		base.SchemaNodeOrder = new string[2] { "f", "sqref" };
	}

	public override string ToString()
	{
		return Cell.Address + ", " + RangeAddress.Address;
	}
}
