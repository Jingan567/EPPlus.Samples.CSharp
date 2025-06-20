using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml;

public sealed class ExcelStyles : XmlHelper
{
	private const string NumberFormatsPath = "d:styleSheet/d:numFmts";

	private const string FontsPath = "d:styleSheet/d:fonts";

	private const string FillsPath = "d:styleSheet/d:fills";

	private const string BordersPath = "d:styleSheet/d:borders";

	private const string CellStyleXfsPath = "d:styleSheet/d:cellStyleXfs";

	private const string CellXfsPath = "d:styleSheet/d:cellXfs";

	private const string CellStylesPath = "d:styleSheet/d:cellStyles";

	private const string dxfsPath = "d:styleSheet/d:dxfs";

	private XmlDocument _styleXml;

	private ExcelWorkbook _wb;

	private XmlNamespaceManager _nameSpaceManager;

	internal int _nextDfxNumFmtID = 164;

	public ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats = new ExcelStyleCollection<ExcelNumberFormatXml>();

	public ExcelStyleCollection<ExcelFontXml> Fonts = new ExcelStyleCollection<ExcelFontXml>();

	public ExcelStyleCollection<ExcelFillXml> Fills = new ExcelStyleCollection<ExcelFillXml>();

	public ExcelStyleCollection<ExcelBorderXml> Borders = new ExcelStyleCollection<ExcelBorderXml>();

	public ExcelStyleCollection<ExcelXfs> CellStyleXfs = new ExcelStyleCollection<ExcelXfs>();

	public ExcelStyleCollection<ExcelXfs> CellXfs = new ExcelStyleCollection<ExcelXfs>();

	public ExcelStyleCollection<ExcelNamedStyleXml> NamedStyles = new ExcelStyleCollection<ExcelNamedStyleXml>();

	public ExcelStyleCollection<ExcelDxfStyleConditionalFormatting> Dxfs = new ExcelStyleCollection<ExcelDxfStyleConditionalFormatting>();

	internal string Id => "";

	internal ExcelStyles(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb)
		: base(NameSpaceManager, xml)
	{
		_styleXml = xml;
		_wb = wb;
		_nameSpaceManager = NameSpaceManager;
		base.SchemaNodeOrder = new string[8] { "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs" };
		LoadFromDocument();
	}

	private void LoadFromDocument()
	{
		ExcelNumberFormatXml.AddBuildIn(base.NameSpaceManager, NumberFormats);
		XmlNode xmlNode = _styleXml.SelectSingleNode("d:styleSheet/d:numFmts", _nameSpaceManager);
		if (xmlNode != null)
		{
			foreach (XmlNode item in xmlNode)
			{
				ExcelNumberFormatXml excelNumberFormatXml = new ExcelNumberFormatXml(_nameSpaceManager, item);
				NumberFormats.Add(excelNumberFormatXml.Id, excelNumberFormatXml);
				if (excelNumberFormatXml.NumFmtId >= NumberFormats.NextId)
				{
					NumberFormats.NextId = excelNumberFormatXml.NumFmtId + 1;
				}
			}
		}
		foreach (XmlNode item2 in _styleXml.SelectSingleNode("d:styleSheet/d:fonts", _nameSpaceManager))
		{
			ExcelFontXml excelFontXml = new ExcelFontXml(_nameSpaceManager, item2);
			Fonts.Add(excelFontXml.Id, excelFontXml);
		}
		foreach (XmlNode item3 in _styleXml.SelectSingleNode("d:styleSheet/d:fills", _nameSpaceManager))
		{
			ExcelFillXml excelFillXml = ((item3.FirstChild == null || !(item3.FirstChild.LocalName == "gradientFill")) ? new ExcelFillXml(_nameSpaceManager, item3) : new ExcelGradientFillXml(_nameSpaceManager, item3));
			Fills.Add(excelFillXml.Id, excelFillXml);
		}
		foreach (XmlNode item4 in _styleXml.SelectSingleNode("d:styleSheet/d:borders", _nameSpaceManager))
		{
			ExcelBorderXml excelBorderXml = new ExcelBorderXml(_nameSpaceManager, item4);
			Borders.Add(excelBorderXml.Id, excelBorderXml);
		}
		XmlNode xmlNode3 = _styleXml.SelectSingleNode("d:styleSheet/d:cellStyleXfs", _nameSpaceManager);
		if (xmlNode3 != null)
		{
			foreach (XmlNode item5 in xmlNode3)
			{
				ExcelXfs excelXfs = new ExcelXfs(_nameSpaceManager, item5, this);
				CellStyleXfs.Add(excelXfs.Id, excelXfs);
			}
		}
		XmlNode xmlNode4 = _styleXml.SelectSingleNode("d:styleSheet/d:cellXfs", _nameSpaceManager);
		for (int i = 0; i < xmlNode4.ChildNodes.Count; i++)
		{
			XmlNode topNode5 = xmlNode4.ChildNodes[i];
			ExcelXfs excelXfs2 = new ExcelXfs(_nameSpaceManager, topNode5, this);
			CellXfs.Add(excelXfs2.Id, excelXfs2);
		}
		XmlNode xmlNode5 = _styleXml.SelectSingleNode("d:styleSheet/d:cellStyles", _nameSpaceManager);
		if (xmlNode5 != null)
		{
			foreach (XmlNode item6 in xmlNode5)
			{
				ExcelNamedStyleXml excelNamedStyleXml = new ExcelNamedStyleXml(_nameSpaceManager, item6, this);
				NamedStyles.Add(excelNamedStyleXml.Name, excelNamedStyleXml);
			}
		}
		XmlNode xmlNode6 = _styleXml.SelectSingleNode("d:styleSheet/d:dxfs", _nameSpaceManager);
		if (xmlNode6 == null)
		{
			return;
		}
		foreach (XmlNode item7 in xmlNode6)
		{
			ExcelDxfStyleConditionalFormatting excelDxfStyleConditionalFormatting = new ExcelDxfStyleConditionalFormatting(_nameSpaceManager, item7, this);
			Dxfs.Add(excelDxfStyleConditionalFormatting.Id, excelDxfStyleConditionalFormatting);
		}
	}

	internal ExcelStyle GetStyleObject(int Id, int PositionID, string Address)
	{
		if (Id < 0)
		{
			Id = 0;
		}
		return new ExcelStyle(this, PropertyChange, PositionID, Address, Id);
	}

	internal int PropertyChange(StyleBase sender, StyleChangeEventArgs e)
	{
		ExcelAddressBase excelAddressBase = new ExcelAddressBase(e.Address);
		ExcelWorksheet excelWorksheet = _wb.Worksheets[e.PositionID];
		Dictionary<int, int> styleCashe = new Dictionary<int, int>();
		lock (excelWorksheet._values)
		{
			SetStyleAddress(sender, e, excelAddressBase, excelWorksheet, ref styleCashe);
			if (excelAddressBase.Addresses != null)
			{
				foreach (ExcelAddress address in excelAddressBase.Addresses)
				{
					SetStyleAddress(sender, e, address, excelWorksheet, ref styleCashe);
				}
			}
		}
		return 0;
	}

	private void SetStyleAddress(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, ref Dictionary<int, int> styleCashe)
	{
		if (address.Start.Column == 0 || address.Start.Row == 0)
		{
			throw new Exception("error address");
		}
		if (address.Start.Row == 1 && address.End.Row == 1048576)
		{
			int col = address.Start.Column;
			int row2 = 0;
			object value = null;
			ExcelColumn excelColumn;
			bool flag;
			if (!ws.ExistsValueInner(0, address.Start.Column, ref value))
			{
				excelColumn = ws.Column(address.Start.Column);
				flag = true;
			}
			else
			{
				excelColumn = (ExcelColumn)value;
				flag = false;
			}
			int columnMax = excelColumn.ColumnMax;
			while (excelColumn.ColumnMin <= address.End.Column)
			{
				if (excelColumn.ColumnMin > columnMax + 1)
				{
					ExcelColumn excelColumn2 = ws.Column(columnMax + 1);
					excelColumn2.ColumnMax = excelColumn.ColumnMin - 1;
					AddNewStyleColumn(sender, e, ws, styleCashe, excelColumn2, excelColumn2.StyleID);
				}
				if (excelColumn.ColumnMax > address.End.Column)
				{
					ws.CopyColumn(excelColumn, address.End.Column + 1, excelColumn.ColumnMax);
					excelColumn.ColumnMax = address.End.Column;
				}
				int styleInner = ws.GetStyleInner(0, excelColumn.ColumnMin);
				AddNewStyleColumn(sender, e, ws, styleCashe, excelColumn, styleInner);
				columnMax = excelColumn.ColumnMax;
				if (!ws._values.NextCell(ref row2, ref col) || row2 > 0)
				{
					if (excelColumn._columnMax != address.End.Column)
					{
						if (flag)
						{
							excelColumn._columnMax = address.End.Column;
							break;
						}
						ExcelColumn excelColumn3 = ws.Column(excelColumn._columnMax + 1);
						excelColumn3.ColumnMax = address.End.Column;
						AddNewStyleColumn(sender, e, ws, styleCashe, excelColumn3, excelColumn3.StyleID);
						excelColumn = excelColumn3;
					}
					break;
				}
				excelColumn = ws.GetValueInner(0, col) as ExcelColumn;
			}
			if (excelColumn._columnMax < address.End.Column)
			{
				ws.Column(excelColumn._columnMax + 1)._columnMax = address.End.Column;
				int styleInner2 = ws.GetStyleInner(0, excelColumn.ColumnMin);
				if (styleCashe.ContainsKey(styleInner2))
				{
					ws.SetStyleInner(0, excelColumn.ColumnMin, styleCashe[styleInner2]);
				}
				else
				{
					int newID = CellXfs[styleInner2].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
					styleCashe.Add(styleInner2, newID);
					ws.SetStyleInner(0, excelColumn.ColumnMin, newID);
				}
				excelColumn._columnMax = address.End.Column;
			}
			CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, address._fromCol, address._toRow, address._toCol);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Column >= address.Start.Column && cellsStoreEnumerator.Column <= address.End.Column && cellsStoreEnumerator.Value._styleId != 0)
				{
					if (styleCashe.ContainsKey(cellsStoreEnumerator.Value._styleId))
					{
						ws.SetStyleInner(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, styleCashe[cellsStoreEnumerator.Value._styleId]);
						continue;
					}
					int newID2 = CellXfs[cellsStoreEnumerator.Value._styleId].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
					styleCashe.Add(cellsStoreEnumerator.Value._styleId, newID2);
					ws.SetStyleInner(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, newID2);
				}
			}
			if (address._fromCol == 1 && address._toCol == 16384)
			{
				return;
			}
			cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, 0, address._toRow, 0);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Value._styleId == 0)
				{
					continue;
				}
				for (int i = address._fromCol; i <= address._toCol; i++)
				{
					if (!ws.ExistsStyleInner(cellsStoreEnumerator.Row, i))
					{
						if (styleCashe.ContainsKey(cellsStoreEnumerator.Value._styleId))
						{
							ws.SetStyleInner(cellsStoreEnumerator.Row, i, styleCashe[cellsStoreEnumerator.Value._styleId]);
							continue;
						}
						int newID3 = CellXfs[cellsStoreEnumerator.Value._styleId].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
						styleCashe.Add(cellsStoreEnumerator.Value._styleId, newID3);
						ws.SetStyleInner(cellsStoreEnumerator.Row, i, newID3);
					}
				}
			}
			return;
		}
		if (address.Start.Column == 1 && address.End.Column == 16384)
		{
			for (int j = address.Start.Row; j <= address.End.Row; j++)
			{
				int num = ws.GetStyleInner(j, 0);
				if (num == 0)
				{
					CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, 16384);
					while (cellsStoreEnumerator2.Next())
					{
						num = cellsStoreEnumerator2.Value._styleId;
						if (num == 0 || !(ws.GetValueInner(cellsStoreEnumerator2.Row, cellsStoreEnumerator2.Column) is ExcelColumn { ColumnMax: <16384, ColumnMin: var k } excelColumn4))
						{
							continue;
						}
						for (; k < excelColumn4.ColumnMax; k++)
						{
							if (!ws.ExistsStyleInner(j, k))
							{
								ws.SetStyleInner(j, k, num);
							}
						}
					}
					ws.SetStyleInner(j, 0, num);
					cellsStoreEnumerator2.Dispose();
				}
				if (styleCashe.ContainsKey(num))
				{
					ws.SetStyleInner(j, 0, styleCashe[num]);
					continue;
				}
				int newID4 = CellXfs[num].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
				styleCashe.Add(num, newID4);
				ws.SetStyleInner(j, 0, newID4);
			}
			CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator3 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);
			while (cellsStoreEnumerator3.Next())
			{
				int styleId = cellsStoreEnumerator3.Value._styleId;
				if (styleId != 0)
				{
					if (styleCashe.ContainsKey(styleId))
					{
						ws.SetStyleInner(cellsStoreEnumerator3.Row, cellsStoreEnumerator3.Column, styleCashe[styleId]);
						continue;
					}
					int newID5 = CellXfs[styleId].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
					styleCashe.Add(styleId, newID5);
					ws.SetStyleInner(cellsStoreEnumerator3.Row, cellsStoreEnumerator3.Column, newID5);
				}
			}
			cellsStoreEnumerator3 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, address._toCol);
			while (cellsStoreEnumerator3.Next())
			{
				if (cellsStoreEnumerator3.Value._styleId == 0)
				{
					continue;
				}
				for (int l = address._fromRow; l <= address._toRow; l++)
				{
					if (!ws.ExistsStyleInner(l, cellsStoreEnumerator3.Column))
					{
						int styleId2 = cellsStoreEnumerator3.Value._styleId;
						if (styleCashe.ContainsKey(styleId2))
						{
							ws.SetStyleInner(l, cellsStoreEnumerator3.Column, styleCashe[styleId2]);
							continue;
						}
						int newID6 = CellXfs[styleId2].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
						styleCashe.Add(styleId2, newID6);
						ws.SetStyleInner(l, cellsStoreEnumerator3.Column, newID6);
					}
				}
			}
			return;
		}
		Dictionary<int, int> tmpCache = styleCashe;
		Dictionary<int, int> rowCache = new Dictionary<int, int>(address.End.Row - address.Start.Row + 1);
		Dictionary<int, ExcelCoreValue> colCache = new Dictionary<int, ExcelCoreValue>(address.End.Column - address.Start.Column + 1);
		ws._values.SetRangeValueSpecial(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, delegate(List<ExcelCoreValue> list, int index, int row, int column, object args)
		{
			int styleId3 = list[index]._styleId;
			if (styleId3 == 0 && !ws.ExistsStyleInner(row, 0, ref styleId3))
			{
				if (!rowCache.ContainsKey(row))
				{
					rowCache.Add(row, ws._values.GetValue(row, 0)._styleId);
				}
				styleId3 = rowCache[row];
				if (styleId3 == 0)
				{
					if (!colCache.ContainsKey(column))
					{
						colCache.Add(column, ws._values.GetValue(0, column));
					}
					styleId3 = colCache[column]._styleId;
					if (styleId3 == 0)
					{
						int row3 = 0;
						int col2 = column;
						if (ws._values.PrevCell(ref row3, ref col2))
						{
							if (!colCache.ContainsKey(col2))
							{
								colCache.Add(col2, ws._values.GetValue(0, col2));
							}
							ExcelCoreValue excelCoreValue = colCache[col2];
							ExcelColumn excelColumn5 = (ExcelColumn)excelCoreValue._value;
							if (excelColumn5 != null && excelColumn5.ColumnMax >= column)
							{
								styleId3 = excelCoreValue._styleId;
							}
						}
					}
				}
			}
			if (tmpCache.ContainsKey(styleId3))
			{
				list[index] = new ExcelCoreValue
				{
					_value = list[index]._value,
					_styleId = tmpCache[styleId3]
				};
			}
			else
			{
				int newID7 = CellXfs[styleId3].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
				tmpCache.Add(styleId3, newID7);
				list[index] = new ExcelCoreValue
				{
					_value = list[index]._value,
					_styleId = newID7
				};
			}
		}, e);
	}

	private void AddNewStyleColumn(StyleBase sender, StyleChangeEventArgs e, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, int s)
	{
		if (styleCashe.ContainsKey(s))
		{
			ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
			return;
		}
		int newID = CellXfs[s].GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
		styleCashe.Add(s, newID);
		ws.SetStyleInner(0, column.ColumnMin, newID);
	}

	internal int GetStyleId(ExcelWorksheet ws, int row, int col)
	{
		int styleId = 0;
		if (ws.ExistsStyleInner(row, col, ref styleId))
		{
			return styleId;
		}
		if (ws.ExistsStyleInner(row, 0, ref styleId))
		{
			return styleId;
		}
		if (ws.ExistsStyleInner(0, col, ref styleId))
		{
			return styleId;
		}
		int row2 = 0;
		int col2 = col;
		if (ws._values.PrevCell(ref row2, ref col2))
		{
			ExcelCoreValue value = ws._values.GetValue(0, col2);
			ExcelColumn excelColumn = (ExcelColumn)value._value;
			if (excelColumn != null && excelColumn.ColumnMax >= col)
			{
				return value._styleId;
			}
			return 0;
		}
		return 0;
	}

	internal int NamedStylePropertyChange(StyleBase sender, StyleChangeEventArgs e)
	{
		int num = NamedStyles.FindIndexByID(e.Address);
		if (num >= 0)
		{
			int newID = CellStyleXfs[NamedStyles[num].StyleXfId].GetNewID(CellStyleXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
			int styleXfId = NamedStyles[num].StyleXfId;
			NamedStyles[num].StyleXfId = newID;
			NamedStyles[num].Style.Index = newID;
			NamedStyles[num].XfId = int.MinValue;
			foreach (ExcelXfs cellXf in CellXfs)
			{
				if (cellXf.XfId == styleXfId)
				{
					cellXf.XfId = newID;
				}
			}
		}
		return 0;
	}

	public ExcelNamedStyleXml CreateNamedStyle(string name)
	{
		return CreateNamedStyle(name, null);
	}

	public ExcelNamedStyleXml CreateNamedStyle(string name, ExcelStyle Template)
	{
		if (_wb.Styles.NamedStyles.ExistsKey(name))
		{
			throw new Exception($"Key {name} already exists in collection");
		}
		ExcelNamedStyleXml excelNamedStyleXml = new ExcelNamedStyleXml(base.NameSpaceManager, this);
		int styleID;
		int positionID;
		ExcelStyles style;
		if (Template == null)
		{
			styleID = 0;
			positionID = -1;
			style = this;
		}
		else if (Template.PositionID < 0 && Template.Styles == this)
		{
			styleID = Template.Index;
			positionID = Template.PositionID;
			style = this;
		}
		else
		{
			styleID = Template.XfId;
			positionID = -1;
			style = Template.Styles;
		}
		int num = CloneStyle(style, styleID, isNamedStyle: true);
		CellStyleXfs[num].XfId = CellStyleXfs.Count - 1;
		int positionID2 = CloneStyle(style, styleID, isNamedStyle: false, allwaysAdd: true);
		CellXfs[positionID2].XfId = num;
		excelNamedStyleXml.Style = new ExcelStyle(this, NamedStylePropertyChange, positionID, name, num);
		excelNamedStyleXml.StyleXfId = num;
		excelNamedStyleXml.Name = name;
		int index = _wb.Styles.NamedStyles.Add(excelNamedStyleXml.Name, excelNamedStyleXml);
		excelNamedStyleXml.Style.SetIndex(index);
		return excelNamedStyleXml;
	}

	public void UpdateXml()
	{
		RemoveUnusedStyles();
		XmlNode xmlNode = _styleXml.SelectSingleNode("d:styleSheet/d:numFmts", _nameSpaceManager);
		if (xmlNode == null)
		{
			CreateNode("d:styleSheet/d:numFmts", insertFirst: true);
			xmlNode = _styleXml.SelectSingleNode("d:styleSheet/d:numFmts", _nameSpaceManager);
		}
		else
		{
			xmlNode.RemoveAll();
		}
		int num = 0;
		int num2 = NamedStyles.FindIndexByID("Normal");
		if (NamedStyles.Count > 0 && num2 >= 0 && NamedStyles[num2].Style.Numberformat.NumFmtID >= 164)
		{
			ExcelNumberFormatXml excelNumberFormatXml = NumberFormats[NumberFormats.FindIndexByID(NamedStyles[num2].Style.Numberformat.Id)];
			xmlNode.AppendChild(excelNumberFormatXml.CreateXmlNode(_styleXml.CreateElement("numFmt", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
			excelNumberFormatXml.newID = num++;
		}
		foreach (ExcelNumberFormatXml numberFormat in NumberFormats)
		{
			if (!numberFormat.BuildIn)
			{
				xmlNode.AppendChild(numberFormat.CreateXmlNode(_styleXml.CreateElement("numFmt", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
				numberFormat.newID = num;
				num++;
			}
		}
		(xmlNode as XmlElement).SetAttribute("count", num.ToString());
		num = 0;
		XmlNode xmlNode2 = _styleXml.SelectSingleNode("d:styleSheet/d:fonts", _nameSpaceManager);
		xmlNode2.RemoveAll();
		if (NamedStyles.Count > 0 && num2 >= 0 && NamedStyles[num2].Style.Font.Index > 0)
		{
			ExcelFontXml excelFontXml = Fonts[NamedStyles[num2].Style.Font.Index];
			xmlNode2.AppendChild(excelFontXml.CreateXmlNode(_styleXml.CreateElement("font", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
			excelFontXml.newID = num++;
		}
		foreach (ExcelFontXml font in Fonts)
		{
			if (font.useCnt > 0)
			{
				xmlNode2.AppendChild(font.CreateXmlNode(_styleXml.CreateElement("font", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
				font.newID = num;
				num++;
			}
		}
		(xmlNode2 as XmlElement).SetAttribute("count", num.ToString());
		num = 0;
		XmlNode xmlNode3 = _styleXml.SelectSingleNode("d:styleSheet/d:fills", _nameSpaceManager);
		xmlNode3.RemoveAll();
		Fills[0].useCnt = 1L;
		Fills[1].useCnt = 1L;
		foreach (ExcelFillXml fill in Fills)
		{
			if (fill.useCnt > 0)
			{
				xmlNode3.AppendChild(fill.CreateXmlNode(_styleXml.CreateElement("fill", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
				fill.newID = num;
				num++;
			}
		}
		(xmlNode3 as XmlElement).SetAttribute("count", num.ToString());
		num = 0;
		XmlNode xmlNode4 = _styleXml.SelectSingleNode("d:styleSheet/d:borders", _nameSpaceManager);
		xmlNode4.RemoveAll();
		Borders[0].useCnt = 1L;
		foreach (ExcelBorderXml border in Borders)
		{
			if (border.useCnt > 0)
			{
				xmlNode4.AppendChild(border.CreateXmlNode(_styleXml.CreateElement("border", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
				border.newID = num;
				num++;
			}
		}
		(xmlNode4 as XmlElement).SetAttribute("count", num.ToString());
		XmlNode xmlNode5 = _styleXml.SelectSingleNode("d:styleSheet/d:cellStyleXfs", _nameSpaceManager);
		if (xmlNode5 == null && NamedStyles.Count > 0)
		{
			CreateNode("d:styleSheet/d:cellStyleXfs");
			xmlNode5 = _styleXml.SelectSingleNode("d:styleSheet/d:cellStyleXfs", _nameSpaceManager);
		}
		if (NamedStyles.Count > 0)
		{
			xmlNode5.RemoveAll();
		}
		num = ((num2 > -1) ? 1 : 0);
		XmlNode xmlNode6 = _styleXml.SelectSingleNode("d:styleSheet/d:cellStyles", _nameSpaceManager);
		xmlNode6?.RemoveAll();
		XmlNode xmlNode7 = _styleXml.SelectSingleNode("d:styleSheet/d:cellXfs", _nameSpaceManager);
		xmlNode7.RemoveAll();
		if (NamedStyles.Count > 0 && num2 >= 0)
		{
			NamedStyles[num2].newID = 0;
			AddNamedStyle(0, xmlNode5, xmlNode7, NamedStyles[num2]);
		}
		foreach (ExcelNamedStyleXml namedStyle in NamedStyles)
		{
			if (!namedStyle.Name.Equals("normal", StringComparison.OrdinalIgnoreCase))
			{
				AddNamedStyle(num++, xmlNode5, xmlNode7, namedStyle);
			}
			else
			{
				namedStyle.newID = 0;
			}
			xmlNode6.AppendChild(namedStyle.CreateXmlNode(_styleXml.CreateElement("cellStyle", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
		}
		if (xmlNode6 != null)
		{
			(xmlNode6 as XmlElement).SetAttribute("count", num.ToString());
		}
		if (xmlNode5 != null)
		{
			(xmlNode5 as XmlElement).SetAttribute("count", num.ToString());
		}
		int num3 = 0;
		foreach (ExcelXfs cellXf in CellXfs)
		{
			if (cellXf.useCnt > 0 && (num2 < 0 || NamedStyles[num2].StyleXfId != num3))
			{
				xmlNode7.AppendChild(cellXf.CreateXmlNode(_styleXml.CreateElement("xf", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
				cellXf.newID = num;
				num++;
			}
			num3++;
		}
		(xmlNode7 as XmlElement).SetAttribute("count", num.ToString());
		XmlNode xmlNode8 = _styleXml.SelectSingleNode("d:styleSheet/d:dxfs", _nameSpaceManager);
		foreach (ExcelWorksheet worksheet in _wb.Worksheets)
		{
			if (worksheet is ExcelChartsheet)
			{
				continue;
			}
			foreach (IExcelConditionalFormattingRule item in (IEnumerable<IExcelConditionalFormattingRule>)worksheet.ConditionalFormatting)
			{
				if (item.Style.HasValue)
				{
					int num4 = Dxfs.FindIndexByID(item.Style.Id);
					if (num4 < 0)
					{
						((ExcelConditionalFormattingRule)item).DxfId = Dxfs.Count;
						Dxfs.Add(item.Style.Id, item.Style);
						XmlElement xmlElement = ((XmlDocument)base.TopNode).CreateElement("d", "dxf", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
						item.Style.CreateNodes(new XmlHelperInstance(base.NameSpaceManager, xmlElement), "");
						xmlNode8.AppendChild(xmlElement);
					}
					else
					{
						((ExcelConditionalFormattingRule)item).DxfId = num4;
					}
				}
			}
		}
		if (xmlNode8 != null)
		{
			(xmlNode8 as XmlElement).SetAttribute("count", Dxfs.Count.ToString());
		}
	}

	private void AddNamedStyle(int id, XmlNode styleXfsNode, XmlNode cellXfsNode, ExcelNamedStyleXml style)
	{
		ExcelXfs excelXfs = CellStyleXfs[style.StyleXfId];
		styleXfsNode.AppendChild(excelXfs.CreateXmlNode(_styleXml.CreateElement("xf", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"), isCellStyleXsf: true));
		excelXfs.newID = id;
		excelXfs.XfId = style.StyleXfId;
		int num = CellXfs.FindIndexByID(excelXfs.Id);
		if (num < 0)
		{
			cellXfsNode.AppendChild(excelXfs.CreateXmlNode(_styleXml.CreateElement("xf", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
		}
		else
		{
			if (id < 0)
			{
				CellXfs[num].XfId = id;
			}
			cellXfsNode.AppendChild(CellXfs[num].CreateXmlNode(_styleXml.CreateElement("xf", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
			CellXfs[num].useCnt = 0L;
			CellXfs[num].newID = id;
		}
		if (style.XfId >= 0)
		{
			style.XfId = CellXfs[style.XfId].newID;
		}
		else
		{
			style.XfId = 0;
		}
	}

	private void RemoveUnusedStyles()
	{
		CellXfs[0].useCnt = 1L;
		foreach (ExcelWorksheet worksheet in _wb.Worksheets)
		{
			CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(worksheet._values);
			while (cellsStoreEnumerator.Next())
			{
				int styleId = cellsStoreEnumerator.Value._styleId;
				if (styleId >= 0)
				{
					CellXfs[styleId].useCnt++;
				}
			}
		}
		foreach (ExcelNamedStyleXml namedStyle in NamedStyles)
		{
			CellStyleXfs[namedStyle.StyleXfId].useCnt++;
		}
		foreach (ExcelXfs cellXf in CellXfs)
		{
			if (cellXf.useCnt > 0)
			{
				if (cellXf.FontId >= 0)
				{
					Fonts[cellXf.FontId].useCnt++;
				}
				if (cellXf.FillId >= 0)
				{
					Fills[cellXf.FillId].useCnt++;
				}
				if (cellXf.BorderId >= 0)
				{
					Borders[cellXf.BorderId].useCnt++;
				}
			}
		}
		foreach (ExcelXfs cellStyleXf in CellStyleXfs)
		{
			if (cellStyleXf.useCnt > 0)
			{
				if (cellStyleXf.FontId >= 0)
				{
					Fonts[cellStyleXf.FontId].useCnt++;
				}
				if (cellStyleXf.FillId >= 0)
				{
					Fills[cellStyleXf.FillId].useCnt++;
				}
				if (cellStyleXf.BorderId >= 0)
				{
					Borders[cellStyleXf.BorderId].useCnt++;
				}
			}
		}
	}

	internal int GetStyleIdFromName(string Name)
	{
		int num = NamedStyles.FindIndexByID(Name);
		if (num >= 0)
		{
			int num2 = NamedStyles[num].XfId;
			if (num2 < 0)
			{
				int styleXfId = NamedStyles[num].StyleXfId;
				ExcelXfs excelXfs = CellStyleXfs[styleXfId].Copy();
				excelXfs.XfId = styleXfId;
				num2 = CellXfs.FindIndexByID(excelXfs.Id);
				if (num2 < 0)
				{
					num2 = CellXfs.Add(excelXfs.Id, excelXfs);
				}
				NamedStyles[num].XfId = num2;
			}
			return num2;
		}
		return 0;
	}

	private int GetXmlNodeInt(XmlNode node)
	{
		if (int.TryParse(GetXmlNode(node), out var result))
		{
			return result;
		}
		return 0;
	}

	private string GetXmlNode(XmlNode node)
	{
		if (node == null)
		{
			return "";
		}
		if (node.Value != null)
		{
			return node.Value;
		}
		return "";
	}

	internal int CloneStyle(ExcelStyles style, int styleID)
	{
		return CloneStyle(style, styleID, isNamedStyle: false, allwaysAdd: false);
	}

	internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle)
	{
		return CloneStyle(style, styleID, isNamedStyle, allwaysAdd: false);
	}

	internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle, bool allwaysAdd)
	{
		lock (style)
		{
			ExcelXfs excelXfs = ((!isNamedStyle) ? style.CellXfs[styleID] : style.CellStyleXfs[styleID]);
			ExcelXfs excelXfs2 = excelXfs.Copy(this);
			if (excelXfs.NumberFormatId > 0)
			{
				string text = string.Empty;
				foreach (ExcelNumberFormatXml numberFormat in style.NumberFormats)
				{
					if (numberFormat.NumFmtId == excelXfs.NumberFormatId)
					{
						text = numberFormat.Format;
						break;
					}
				}
				if (!string.IsNullOrEmpty(text))
				{
					int num = NumberFormats.FindIndexByID(text);
					if (num < 0)
					{
						ExcelNumberFormatXml excelNumberFormatXml = new ExcelNumberFormatXml(base.NameSpaceManager)
						{
							Format = text,
							NumFmtId = NumberFormats.NextId++
						};
						NumberFormats.Add(text, excelNumberFormatXml);
						excelXfs2.NumberFormatId = excelNumberFormatXml.NumFmtId;
					}
					else
					{
						excelXfs2.NumberFormatId = NumberFormats[num].NumFmtId;
					}
				}
			}
			if (excelXfs.FontId > -1)
			{
				int num2 = Fonts.FindIndexByID(excelXfs.Font.Id);
				if (num2 < 0)
				{
					ExcelFontXml item = style.Fonts[excelXfs.FontId].Copy();
					num2 = Fonts.Add(excelXfs.Font.Id, item);
				}
				excelXfs2.FontId = num2;
			}
			if (excelXfs.BorderId > -1)
			{
				int num3 = Borders.FindIndexByID(excelXfs.Border.Id);
				if (num3 < 0)
				{
					ExcelBorderXml item2 = style.Borders[excelXfs.BorderId].Copy();
					num3 = Borders.Add(excelXfs.Border.Id, item2);
				}
				excelXfs2.BorderId = num3;
			}
			if (excelXfs.FillId > -1)
			{
				int num4 = Fills.FindIndexByID(excelXfs.Fill.Id);
				if (num4 < 0)
				{
					ExcelFillXml item3 = style.Fills[excelXfs.FillId].Copy();
					num4 = Fills.Add(excelXfs.Fill.Id, item3);
				}
				excelXfs2.FillId = num4;
			}
			if (excelXfs.XfId > 0)
			{
				string id = style.CellStyleXfs[excelXfs.XfId].Id;
				int num5 = CellStyleXfs.FindIndexByID(id);
				if (num5 >= 0)
				{
					excelXfs2.XfId = num5;
				}
				else if (style._wb != _wb && !allwaysAdd)
				{
					Dictionary<int, ExcelNamedStyleXml> dictionary = style.NamedStyles.ToDictionary((ExcelNamedStyleXml d) => d.StyleXfId);
					if (dictionary.ContainsKey(excelXfs.XfId))
					{
						ExcelNamedStyleXml excelNamedStyleXml = dictionary[excelXfs.XfId];
						if (NamedStyles.ExistsKey(excelNamedStyleXml.Name))
						{
							excelXfs2.XfId = NamedStyles.FindIndexByID(excelNamedStyleXml.Name);
						}
						else
						{
							CreateNamedStyle(excelNamedStyleXml.Name, excelNamedStyleXml.Style);
							excelXfs2.XfId = NamedStyles.Count - 1;
						}
					}
				}
			}
			int num6;
			if (isNamedStyle)
			{
				num6 = CellStyleXfs.Add(excelXfs2.Id, excelXfs2);
			}
			else if (allwaysAdd)
			{
				num6 = CellXfs.Add(excelXfs2.Id, excelXfs2);
			}
			else
			{
				num6 = CellXfs.FindIndexByID(excelXfs2.Id);
				if (num6 < 0)
				{
					num6 = CellXfs.Add(excelXfs2.Id, excelXfs2);
				}
			}
			return num6;
		}
	}
}
