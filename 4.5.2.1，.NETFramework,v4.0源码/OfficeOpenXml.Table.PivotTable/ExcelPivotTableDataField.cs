using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableDataField : XmlHelper
{
	public ExcelPivotTableField Field { get; private set; }

	public int Index
	{
		get
		{
			return GetXmlNodeInt("@fld");
		}
		internal set
		{
			SetXmlNodeString("@fld", value.ToString());
		}
	}

	public string Name
	{
		get
		{
			return GetXmlNodeString("@name");
		}
		set
		{
			if (Field._table.DataFields.ExistsDfName(value, this))
			{
				throw new InvalidOperationException("Duplicate datafield name");
			}
			SetXmlNodeString("@name", value);
		}
	}

	public int BaseField
	{
		get
		{
			return GetXmlNodeInt("@baseField");
		}
		set
		{
			SetXmlNodeString("@baseField", value.ToString());
		}
	}

	public int BaseItem
	{
		get
		{
			return GetXmlNodeInt("@baseItem");
		}
		set
		{
			SetXmlNodeString("@baseItem", value.ToString());
		}
	}

	internal int NumFmtId
	{
		get
		{
			return GetXmlNodeInt("@numFmtId");
		}
		set
		{
			SetXmlNodeString("@numFmtId", value.ToString());
		}
	}

	public string Format
	{
		get
		{
			foreach (ExcelNumberFormatXml numberFormat in Field._table.WorkSheet.Workbook.Styles.NumberFormats)
			{
				if (numberFormat.NumFmtId == NumFmtId)
				{
					return numberFormat.Format;
				}
			}
			return Field._table.WorkSheet.Workbook.Styles.NumberFormats[0].Format;
		}
		set
		{
			ExcelStyles styles = Field._table.WorkSheet.Workbook.Styles;
			ExcelNumberFormatXml obj = null;
			if (!styles.NumberFormats.FindByID(value, ref obj))
			{
				obj = new ExcelNumberFormatXml(base.NameSpaceManager)
				{
					Format = value,
					NumFmtId = styles.NumberFormats.NextId++
				};
				styles.NumberFormats.Add(value, obj);
			}
			NumFmtId = obj.NumFmtId;
		}
	}

	public DataFieldFunctions Function
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@subtotal");
			if (xmlNodeString == "")
			{
				return DataFieldFunctions.None;
			}
			return (DataFieldFunctions)Enum.Parse(typeof(DataFieldFunctions), xmlNodeString, ignoreCase: true);
		}
		set
		{
			string value2;
			switch (value)
			{
			case DataFieldFunctions.None:
				DeleteNode("@subtotal");
				return;
			case DataFieldFunctions.CountNums:
				value2 = "CountNums";
				break;
			case DataFieldFunctions.StdDev:
				value2 = "stdDev";
				break;
			case DataFieldFunctions.StdDevP:
				value2 = "stdDevP";
				break;
			default:
				value2 = value.ToString().ToLower(CultureInfo.InvariantCulture);
				break;
			}
			SetXmlNodeString("@subtotal", value2);
		}
	}

	internal ExcelPivotTableDataField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field)
		: base(ns, topNode)
	{
		if (topNode.Attributes.Count == 0)
		{
			Index = field.Index;
			BaseField = 0;
			BaseItem = 0;
		}
		Field = field;
	}
}
