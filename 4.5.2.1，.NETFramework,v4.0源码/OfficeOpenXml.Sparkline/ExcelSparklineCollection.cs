using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Sparkline;

public class ExcelSparklineCollection : IEnumerable<ExcelSparkline>, IEnumerable
{
	private ExcelSparklineGroup _slg;

	private List<ExcelSparkline> _lst;

	private const string _topPath = "x14:sparklines/x14:sparkline";

	public int Count => _lst.Count;

	public ExcelSparkline this[int index] => _lst[index];

	internal ExcelSparklineCollection(ExcelSparklineGroup slg)
	{
		_slg = slg;
		_lst = new List<ExcelSparkline>();
		LoadSparklines();
	}

	private void LoadSparklines()
	{
		foreach (XmlElement item in _slg.TopNode.SelectNodes("x14:sparklines/x14:sparkline", _slg.NameSpaceManager))
		{
			_lst.Add(new ExcelSparkline(_slg.NameSpaceManager, item));
		}
	}

	public IEnumerator<ExcelSparkline> GetEnumerator()
	{
		return _lst.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _lst.GetEnumerator();
	}

	internal void Add(ExcelCellAddress cell, string worksheetName, ExcelAddressBase sqref)
	{
		XmlElement xmlElement = _slg.TopNode.OwnerDocument.CreateElement("x14", "sparkline", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		XmlNode xmlNode = _slg.TopNode.SelectSingleNode("x14:sparklines", _slg.NameSpaceManager);
		xmlNode.AppendChild(xmlElement);
		_slg.TopNode.AppendChild(xmlNode);
		ExcelSparkline excelSparkline = new ExcelSparkline(_slg.NameSpaceManager, xmlElement);
		excelSparkline.Cell = cell;
		excelSparkline.RangeAddress = sqref;
		_lst.Add(excelSparkline);
	}
}
