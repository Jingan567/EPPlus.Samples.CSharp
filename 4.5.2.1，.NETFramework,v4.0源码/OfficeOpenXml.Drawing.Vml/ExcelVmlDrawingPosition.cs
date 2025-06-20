using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

public class ExcelVmlDrawingPosition : XmlHelper
{
	private int _startPos;

	public int Row
	{
		get
		{
			return GetNumber(2);
		}
		set
		{
			SetNumber(2, value);
		}
	}

	public int RowOffset
	{
		get
		{
			return GetNumber(3);
		}
		set
		{
			SetNumber(3, value);
		}
	}

	public int Column
	{
		get
		{
			return GetNumber(0);
		}
		set
		{
			SetNumber(0, value);
		}
	}

	public int ColumnOffset
	{
		get
		{
			return GetNumber(1);
		}
		set
		{
			SetNumber(1, value);
		}
	}

	internal ExcelVmlDrawingPosition(XmlNamespaceManager ns, XmlNode topNode, int startPos)
		: base(ns, topNode)
	{
		_startPos = startPos;
	}

	private void SetNumber(int pos, int value)
	{
		string[] array = GetXmlNodeString("x:Anchor").Split(',');
		if (array.Length == 8)
		{
			array[_startPos + pos] = value.ToString();
			SetXmlNodeString("x:Anchor", string.Join(",", array));
			return;
		}
		throw new Exception("Anchor element is invalid in vmlDrawing");
	}

	private int GetNumber(int pos)
	{
		string[] array = GetXmlNodeString("x:Anchor").Split(',');
		if (array.Length == 8 && int.TryParse(array[_startPos + pos], NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		throw new Exception("Anchor element is invalid in vmlDrawing");
	}
}
