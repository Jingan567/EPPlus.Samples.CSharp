using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style;

public class ExcelRichTextCollection : XmlHelper, IEnumerable<ExcelRichText>, IEnumerable
{
	private List<ExcelRichText> _list = new List<ExcelRichText>();

	private ExcelRangeBase _cells;

	public ExcelRichText this[int Index]
	{
		get
		{
			ExcelRichText excelRichText = _list[Index];
			if (_cells != null)
			{
				excelRichText.SetCallback(UpdateCells);
			}
			return excelRichText;
		}
	}

	public int Count => _list.Count;

	public string Text
	{
		get
		{
			StringBuilder stringBuilder = new StringBuilder();
			foreach (ExcelRichText item in _list)
			{
				stringBuilder.Append(item.Text);
			}
			return stringBuilder.ToString();
		}
		set
		{
			if (Count == 0)
			{
				Add(value);
				return;
			}
			this[0].Text = value;
			for (int i = 1; i < Count; i++)
			{
				RemoveAt(i);
			}
		}
	}

	internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode topNode)
		: base(ns, topNode)
	{
		XmlNodeList xmlNodeList = topNode.SelectNodes("d:r", base.NameSpaceManager);
		if (xmlNodeList == null)
		{
			return;
		}
		foreach (XmlNode item in xmlNodeList)
		{
			_list.Add(new ExcelRichText(ns, item, this));
		}
	}

	internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode topNode, ExcelRangeBase cells)
		: this(ns, topNode)
	{
		_cells = cells;
	}

	public ExcelRichText Add(string Text)
	{
		return Insert(_list.Count, Text);
	}

	public ExcelRichText Insert(int index, string text)
	{
		ConvertRichtext();
		XmlDocument xmlDocument = ((!(base.TopNode is XmlDocument)) ? base.TopNode.OwnerDocument : (base.TopNode as XmlDocument));
		XmlElement xmlElement = xmlDocument.CreateElement("d", "r", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		if (index < _list.Count)
		{
			base.TopNode.InsertBefore(xmlElement, base.TopNode.ChildNodes[index]);
		}
		else
		{
			base.TopNode.AppendChild(xmlElement);
		}
		ExcelRichText excelRichText = new ExcelRichText(base.NameSpaceManager, xmlElement, this);
		if (_list.Count > 0)
		{
			ExcelRichText excelRichText2 = _list[(index < _list.Count) ? index : (_list.Count - 1)];
			excelRichText.FontName = excelRichText2.FontName;
			excelRichText.Size = excelRichText2.Size;
			if (excelRichText2.Color.IsEmpty)
			{
				excelRichText.Color = Color.Black;
			}
			else
			{
				excelRichText.Color = excelRichText2.Color;
			}
			excelRichText.PreserveSpace = excelRichText.PreserveSpace;
			excelRichText.Bold = excelRichText2.Bold;
			excelRichText.Italic = excelRichText2.Italic;
			excelRichText.UnderLine = excelRichText2.UnderLine;
		}
		else if (_cells == null)
		{
			excelRichText.FontName = "Calibri";
			excelRichText.Size = 11f;
		}
		else
		{
			ExcelStyle style = _cells.Offset(0, 0).Style;
			excelRichText.FontName = style.Font.Name;
			excelRichText.Size = style.Font.Size;
			excelRichText.Bold = style.Font.Bold;
			excelRichText.Italic = style.Font.Italic;
			_cells.IsRichText = true;
		}
		excelRichText.Text = text;
		excelRichText.PreserveSpace = true;
		if (_cells != null)
		{
			excelRichText.SetCallback(UpdateCells);
			UpdateCells();
		}
		_list.Insert(index, excelRichText);
		return excelRichText;
	}

	internal void ConvertRichtext()
	{
		if (_cells == null)
		{
			return;
		}
		bool flagValue = _cells.Worksheet._flags.GetFlagValue(_cells._fromRow, _cells._fromCol, CellFlags.RichText);
		if (Count == 1 && !flagValue)
		{
			_cells.Worksheet._flags.SetFlagValue(_cells._fromRow, _cells._fromCol, value: true, CellFlags.RichText);
			int styleInner = _cells.Worksheet.GetStyleInner(_cells._fromRow, _cells._fromCol);
			ExcelFont font = _cells.Worksheet.Workbook.Styles.GetStyleObject(styleInner, _cells.Worksheet.PositionID, ExcelCellBase.GetAddress(_cells._fromRow, _cells._fromCol)).Font;
			this[0].PreserveSpace = true;
			this[0].Bold = font.Bold;
			this[0].FontName = font.Name;
			this[0].Italic = font.Italic;
			this[0].Size = font.Size;
			this[0].UnderLine = font.UnderLine;
			if (font.Color.Rgb != "" && int.TryParse(font.Color.Rgb, NumberStyles.HexNumber, null, out var result))
			{
				this[0].Color = Color.FromArgb(result);
			}
		}
	}

	internal void UpdateCells()
	{
		_cells.SetValueRichText(base.TopNode.InnerXml);
	}

	public void Clear()
	{
		_list.Clear();
		base.TopNode.RemoveAll();
		UpdateCells();
		if (_cells != null)
		{
			_cells.IsRichText = false;
		}
	}

	public void RemoveAt(int Index)
	{
		base.TopNode.RemoveChild(_list[Index].TopNode);
		_list.RemoveAt(Index);
		if (_cells != null && _list.Count == 0)
		{
			_cells.IsRichText = false;
		}
	}

	public void Remove(ExcelRichText Item)
	{
		base.TopNode.RemoveChild(Item.TopNode);
		_list.Remove(Item);
		if (_cells != null && _list.Count == 0)
		{
			_cells.IsRichText = false;
		}
	}

	IEnumerator<ExcelRichText> IEnumerable<ExcelRichText>.GetEnumerator()
	{
		return _list.Select(delegate(ExcelRichText x)
		{
			x.SetCallback(UpdateCells);
			return x;
		}).GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.Select(delegate(ExcelRichText x)
		{
			x.SetCallback(UpdateCells);
			return x;
		}).GetEnumerator();
	}
}
