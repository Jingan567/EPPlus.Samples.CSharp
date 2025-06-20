using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Sparkline;

public class ExcelSparklineGroupCollection : IEnumerable<ExcelSparklineGroup>, IEnumerable
{
	private ExcelWorksheet _ws;

	private List<ExcelSparklineGroup> _lst;

	private const string _extPath = "/d:worksheet/d:extLst/d:ext";

	private const string _searchPath = "[@uri='{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']";

	private const string _topSearchPath = "/d:worksheet/d:extLst/d:ext[@uri='{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']/x14:sparklineGroups";

	private const string _topPath = "/d:worksheet/d:extLst/d:ext/x14:sparklineGroups";

	public int Count => _lst.Count;

	public ExcelSparklineGroup this[int index] => _lst[index];

	internal ExcelSparklineGroupCollection(ExcelWorksheet ws)
	{
		_ws = ws;
		_lst = new List<ExcelSparklineGroup>();
		LoadSparklines();
	}

	public ExcelSparklineGroup Add(eSparklineType type, ExcelAddressBase locationRange, ExcelAddressBase dataRange)
	{
		if (locationRange.Rows == 1)
		{
			if (locationRange.Columns == dataRange.Rows)
			{
				return AddGroup(type, locationRange, dataRange, isRows: true);
			}
			if (locationRange.Columns == dataRange.Columns)
			{
				return AddGroup(type, locationRange, dataRange, isRows: false);
			}
			throw new ArgumentException("dataRange is not valid. dataRange columns or rows must match number of rows in locationRange");
		}
		if (locationRange.Columns == 1)
		{
			if (locationRange.Rows == dataRange.Columns)
			{
				return AddGroup(type, locationRange, dataRange, isRows: false);
			}
			if (locationRange.Rows == dataRange.Rows)
			{
				return AddGroup(type, locationRange, dataRange, isRows: true);
			}
			throw new ArgumentException("dataRange is not valid. dataRange columns or rows must match number of columns in locationRange");
		}
		throw new ArgumentException("locationRange is not valid. Range must be one Column or Row only");
	}

	private ExcelSparklineGroup AddGroup(eSparklineType type, ExcelAddressBase locationRange, ExcelAddressBase dataRange, bool isRows)
	{
		ExcelSparklineGroup excelSparklineGroup = NewSparklineGroup();
		excelSparklineGroup.Type = type;
		int num = locationRange._fromRow;
		int num2 = locationRange._fromCol;
		int num3 = dataRange._fromRow;
		int num4 = dataRange._fromCol;
		int num5 = (isRows ? dataRange._fromRow : dataRange._toRow);
		int num6 = (isRows ? dataRange._toCol : dataRange._fromCol);
		int num7 = ((locationRange._fromRow == locationRange._toRow) ? (locationRange._toCol - locationRange._fromCol) : (locationRange._toRow - locationRange._fromRow)) + 1;
		int num8 = 0;
		while (num8 < num7)
		{
			ExcelCellAddress cell = new ExcelCellAddress(num, num2);
			ExcelAddressBase sqref = new ExcelAddressBase(dataRange.WorkSheet, num3, num4, num5, num6);
			excelSparklineGroup.Sparklines.Add(cell, dataRange.WorkSheet, sqref);
			num8++;
			if (locationRange._fromRow == locationRange._toRow)
			{
				num2++;
			}
			else
			{
				num++;
			}
			if (isRows)
			{
				num3++;
				num5++;
			}
			else
			{
				num4++;
				num6++;
			}
		}
		excelSparklineGroup.ColorSeries.Rgb = "FF376092";
		excelSparklineGroup.ColorNegative.Rgb = "FFD00000";
		excelSparklineGroup.ColorMarkers.Rgb = "FFD00000";
		excelSparklineGroup.ColorAxis.Rgb = "FF000000";
		excelSparklineGroup.ColorFirst.Rgb = "FFD00000";
		excelSparklineGroup.ColorLast.Rgb = "FFD00000";
		excelSparklineGroup.ColorHigh.Rgb = "FFD00000";
		excelSparklineGroup.ColorLow.Rgb = "FFD00000";
		return excelSparklineGroup;
	}

	private ExcelSparklineGroup NewSparklineGroup()
	{
		XmlHelperInstance xmlHelperInstance = new XmlHelperInstance(_ws.NameSpaceManager, _ws.WorksheetXml);
		if (!xmlHelperInstance.ExistNode("/d:worksheet/d:extLst/d:ext[@uri='{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']"))
		{
			XmlNode xmlNode = xmlHelperInstance.CreateNode("/d:worksheet/d:extLst/d:ext", insertFirst: false, addNew: true);
			if (xmlNode.Attributes["uri"] == null)
			{
				((XmlElement)xmlNode).SetAttribute("uri", "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}");
			}
		}
		XmlNode xmlNode2 = xmlHelperInstance.CreateNode("/d:worksheet/d:extLst/d:ext[@uri='{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']/x14:sparklineGroups");
		XmlElement xmlElement = _ws.WorksheetXml.CreateElement("x14", "sparklineGroup", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		xmlElement.SetAttribute("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main");
		xmlElement.SetAttribute("uid", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2", $"{{{Guid.NewGuid().ToString()}}}");
		xmlNode2.AppendChild(xmlElement);
		xmlElement.AppendChild(xmlElement.OwnerDocument.CreateElement("x14", "sparklines", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"));
		return new ExcelSparklineGroup(_ws.NameSpaceManager, xmlElement, _ws);
	}

	private void LoadSparklines()
	{
		foreach (XmlElement item in _ws.WorksheetXml.SelectNodes("/d:worksheet/d:extLst/d:ext/x14:sparklineGroups/x14:sparklineGroup", _ws.NameSpaceManager))
		{
			_lst.Add(new ExcelSparklineGroup(_ws.NameSpaceManager, item, _ws));
		}
	}

	public IEnumerator<ExcelSparklineGroup> GetEnumerator()
	{
		return _lst.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _lst.GetEnumerator();
	}

	public void RemoveAt(int index)
	{
		Remove(_lst[index]);
	}

	public void Remove(ExcelSparklineGroup sparklineGroup)
	{
		sparklineGroup.TopNode.ParentNode.RemoveChild(sparklineGroup.TopNode);
		_lst.Remove(sparklineGroup);
	}
}
