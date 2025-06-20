using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

internal class ExcelVmlDrawingCommentCollection : ExcelVmlDrawingBaseCollection, IEnumerable
{
	internal RangeCollection _drawings;

	private int _nextID;

	internal ExcelVmlDrawingBase this[ulong rangeID] => _drawings[rangeID] as ExcelVmlDrawingComment;

	internal int Count => _drawings.Count;

	internal ExcelVmlDrawingCommentCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
		: base(pck, ws, uri)
	{
		if (uri == null)
		{
			base.VmlDrawingXml.LoadXml(CreateVmlDrawings());
			_drawings = new RangeCollection(new List<IRangeID>());
		}
		else
		{
			AddDrawingsFromXml(ws);
		}
	}

	protected void AddDrawingsFromXml(ExcelWorksheet ws)
	{
		XmlNodeList xmlNodeList = base.VmlDrawingXml.SelectNodes("//v:shape", base.NameSpaceManager);
		List<IRangeID> list = new List<IRangeID>();
		foreach (XmlNode item in xmlNodeList)
		{
			XmlNode xmlNode2 = item.SelectSingleNode("x:ClientData/x:Row", base.NameSpaceManager);
			XmlNode xmlNode3 = item.SelectSingleNode("x:ClientData/x:Column", base.NameSpaceManager);
			if (xmlNode2 != null && xmlNode3 != null)
			{
				int row = int.Parse(xmlNode2.InnerText) + 1;
				int col = int.Parse(xmlNode3.InnerText) + 1;
				list.Add(new ExcelVmlDrawingComment(item, ws.Cells[row, col], base.NameSpaceManager));
			}
			else
			{
				list.Add(new ExcelVmlDrawingComment(item, ws.Cells[1, 1], base.NameSpaceManager));
			}
		}
		list.Sort((IRangeID r1, IRangeID r2) => (r1.RangeID >= r2.RangeID) ? ((r1.RangeID > r2.RangeID) ? 1 : 0) : (-1));
		_drawings = new RangeCollection(list);
	}

	private string CreateVmlDrawings()
	{
		return string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">", "urn:schemas-microsoft-com:vml", "urn:schemas-microsoft-com:office:office", "urn:schemas-microsoft-com:office:excel") + "<o:shapelayout v:ext=\"edit\">", "<o:idmap v:ext=\"edit\" data=\"1\"/>"), "</o:shapelayout>"), "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">"), "<v:stroke joinstyle=\"miter\" />"), "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />"), "</v:shapetype>"), "</xml>");
	}

	internal ExcelVmlDrawingComment Add(ExcelRangeBase cell)
	{
		ExcelVmlDrawingComment excelVmlDrawingComment = new ExcelVmlDrawingComment(AddDrawing(cell), cell, base.NameSpaceManager);
		_drawings.Add(excelVmlDrawingComment);
		return excelVmlDrawingComment;
	}

	private XmlNode AddDrawing(ExcelRangeBase cell)
	{
		int row = cell.Start.Row;
		int column = cell.Start.Column;
		XmlElement xmlElement = base.VmlDrawingXml.CreateElement("v", "shape", "urn:schemas-microsoft-com:vml");
		ulong cellID = ExcelCellBase.GetCellID(cell.Worksheet.SheetID, cell._fromRow, cell._fromCol);
		int num = _drawings.IndexOf(cellID);
		if (num < 0 && ~num < _drawings.Count)
		{
			num = ~num;
			ExcelVmlDrawingBase excelVmlDrawingBase = _drawings[num] as ExcelVmlDrawingBase;
			excelVmlDrawingBase.TopNode.ParentNode.InsertBefore(xmlElement, excelVmlDrawingBase.TopNode);
		}
		else
		{
			base.VmlDrawingXml.DocumentElement.AppendChild(xmlElement);
		}
		xmlElement.SetAttribute("id", GetNewId());
		xmlElement.SetAttribute("type", "#_x0000_t202");
		xmlElement.SetAttribute("style", "position:absolute;z-index:1; visibility:hidden");
		xmlElement.SetAttribute("fillcolor", "#ffffe1");
		xmlElement.SetAttribute("insetmode", "urn:schemas-microsoft-com:office:office", "auto");
		string text = "<v:fill color2=\"#ffffe1\" />";
		text += "<v:shadow on=\"t\" color=\"black\" obscured=\"t\" />";
		text += "<v:path o:connecttype=\"none\" />";
		text += "<v:textbox style=\"mso-direction-alt:auto\">";
		text += "<div style=\"text-align:left\" />";
		text += "</v:textbox>";
		text += "<x:ClientData ObjectType=\"Note\">";
		text += "<x:MoveWithCells />";
		text += "<x:SizeWithCells />";
		text += $"<x:Anchor>{column}, 15, {row - 1}, 2, {column + 2}, 31, {row + 3}, 1</x:Anchor>";
		text += "<x:AutoFill>False</x:AutoFill>";
		text += $"<x:Row>{row - 1}</x:Row>";
		text += $"<x:Column>{column - 1}</x:Column>";
		text += "</x:ClientData>";
		xmlElement.InnerXml = text;
		return xmlElement;
	}

	internal string GetNewId()
	{
		if (_nextID == 0)
		{
			{
				IEnumerator enumerator = GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						ExcelVmlDrawingComment excelVmlDrawingComment = (ExcelVmlDrawingComment)enumerator.Current;
						if (excelVmlDrawingComment.Id.Length > 3 && excelVmlDrawingComment.Id.StartsWith("vml") && int.TryParse(excelVmlDrawingComment.Id.Substring(3, excelVmlDrawingComment.Id.Length - 3), NumberStyles.Any, CultureInfo.InvariantCulture, out var result) && result > _nextID)
						{
							_nextID = result;
						}
					}
				}
				finally
				{
					IDisposable disposable = enumerator as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
		}
		_nextID++;
		return "vml" + _nextID;
	}

	internal bool ContainsKey(ulong rangeID)
	{
		return _drawings.ContainsKey(rangeID);
	}

	public IEnumerator GetEnumerator()
	{
		return _drawings;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _drawings;
	}
}
