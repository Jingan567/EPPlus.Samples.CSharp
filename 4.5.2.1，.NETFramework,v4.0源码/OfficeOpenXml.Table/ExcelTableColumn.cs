using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table;

public class ExcelTableColumn : XmlHelper
{
	internal ExcelTable _tbl;

	private const string TOTALSROWFORMULA_PATH = "d:totalsRowFormula";

	private const string DATACELLSTYLE_PATH = "@dataCellStyle";

	private const string CALCULATEDCOLUMNFORMULA_PATH = "d:calculatedColumnFormula";

	public int Id
	{
		get
		{
			return GetXmlNodeInt("@id");
		}
		set
		{
			SetXmlNodeString("@id", value.ToString());
		}
	}

	public int Position { get; private set; }

	public string Name
	{
		get
		{
			string text = GetXmlNodeString("@name");
			if (string.IsNullOrEmpty(text))
			{
				text = ((!_tbl.ShowHeader) ? ("Column" + (Position + 1)) : ConvertUtil.ExcelDecodeString(_tbl.WorkSheet.GetValue<string>(_tbl.Address._fromRow, _tbl.Address._fromCol + Position)));
			}
			return text;
		}
		set
		{
			string value2 = ConvertUtil.ExcelEncodeString(value);
			SetXmlNodeString("@name", value2);
			if (_tbl.ShowHeader)
			{
				_tbl.WorkSheet.SetValue(_tbl.Address._fromRow, _tbl.Address._fromCol + Position, value);
			}
			_tbl.WorkSheet.SetTableTotalFunction(_tbl, this);
		}
	}

	public string TotalsRowLabel
	{
		get
		{
			return GetXmlNodeString("@totalsRowLabel");
		}
		set
		{
			SetXmlNodeString("@totalsRowLabel", value);
		}
	}

	public RowFunctions TotalsRowFunction
	{
		get
		{
			if (GetXmlNodeString("@totalsRowFunction") == "")
			{
				return RowFunctions.None;
			}
			return (RowFunctions)Enum.Parse(typeof(RowFunctions), GetXmlNodeString("@totalsRowFunction"), ignoreCase: true);
		}
		set
		{
			if (value == RowFunctions.Custom)
			{
				throw new Exception("Use the TotalsRowFormula-property to set a custom table formula");
			}
			string text = value.ToString();
			text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
			SetXmlNodeString("@totalsRowFunction", text);
			_tbl.WorkSheet.SetTableTotalFunction(_tbl, this);
		}
	}

	public string TotalsRowFormula
	{
		get
		{
			return GetXmlNodeString("d:totalsRowFormula");
		}
		set
		{
			if (value.StartsWith("="))
			{
				value = value.Substring(1, value.Length - 1);
			}
			SetXmlNodeString("@totalsRowFunction", "custom");
			SetXmlNodeString("d:totalsRowFormula", value);
			_tbl.WorkSheet.SetTableTotalFunction(_tbl, this);
		}
	}

	public string DataCellStyleName
	{
		get
		{
			return GetXmlNodeString("@dataCellStyle");
		}
		set
		{
			if (_tbl.WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
			{
				throw new Exception($"Named style {value} does not exist.");
			}
			SetXmlNodeString(base.TopNode, "@dataCellStyle", value, removeIfBlank: true);
			int num = _tbl.Address._fromRow + (_tbl.ShowHeader ? 1 : 0);
			int num2 = _tbl.Address._toRow - (_tbl.ShowTotal ? 1 : 0);
			int num3 = _tbl.Address._fromCol + Position;
			if (num <= num2)
			{
				_tbl.WorkSheet.Cells[num, num3, num2, num3].StyleName = value;
			}
		}
	}

	public string CalculatedColumnFormula
	{
		get
		{
			return GetXmlNodeString("d:calculatedColumnFormula");
		}
		set
		{
			if (value.StartsWith("="))
			{
				value = value.Substring(1, value.Length - 1);
			}
			SetXmlNodeString("d:calculatedColumnFormula", value);
		}
	}

	internal ExcelTableColumn(XmlNamespaceManager ns, XmlNode topNode, ExcelTable tbl, int pos)
		: base(ns, topNode)
	{
		_tbl = tbl;
		Position = pos;
	}
}
