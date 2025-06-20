using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public class ExcelCommentCollection : IEnumerable, IDisposable
{
	private List<ExcelComment> _list = new List<ExcelComment>();

	public XmlDocument CommentXml { get; set; }

	internal Uri Uri { get; set; }

	internal string RelId { get; set; }

	internal XmlNamespaceManager NameSpaceManager { get; set; }

	internal ZipPackagePart Part { get; set; }

	public ExcelWorksheet Worksheet { get; set; }

	public int Count => _list.Count;

	public ExcelComment this[int Index]
	{
		get
		{
			if (Index < 0 || Index >= _list.Count)
			{
				throw new ArgumentOutOfRangeException("Comment index out of range");
			}
			return _list[Index];
		}
	}

	public ExcelComment this[ExcelCellAddress cell]
	{
		get
		{
			int value = -1;
			if (Worksheet._commentsStore.Exists(cell.Row, cell.Column, ref value))
			{
				return _list[value];
			}
			return null;
		}
	}

	internal ExcelCommentCollection(ExcelPackage pck, ExcelWorksheet ws, XmlNamespaceManager ns)
	{
		CommentXml = new XmlDocument();
		CommentXml.PreserveWhitespace = false;
		NameSpaceManager = ns;
		Worksheet = ws;
		CreateXml(pck);
		AddCommentsFromXml();
	}

	private void CreateXml(ExcelPackage pck)
	{
		ZipPackageRelationshipCollection relationshipsByType = Worksheet.Part.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
		bool flag = false;
		CommentXml = new XmlDocument();
		foreach (ZipPackageRelationship item in relationshipsByType)
		{
			Uri = UriHelper.ResolvePartUri(item.SourceUri, item.TargetUri);
			Part = pck.Package.GetPart(Uri);
			XmlHelper.LoadXmlSafe(CommentXml, Part.GetStream());
			RelId = item.Id;
			flag = true;
		}
		if (!flag)
		{
			CommentXml.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors /><commentList /></comments>");
			Uri = null;
		}
	}

	private void AddCommentsFromXml()
	{
		foreach (XmlElement item in CommentXml.SelectNodes("//d:commentList/d:comment", NameSpaceManager))
		{
			ExcelComment excelComment = new ExcelComment(NameSpaceManager, item, new ExcelRangeBase(Worksheet, item.GetAttribute("ref")));
			_list.Add(excelComment);
			Worksheet._commentsStore.SetValue(excelComment.Range._fromRow, excelComment.Range._fromCol, _list.Count - 1);
		}
	}

	public ExcelComment Add(ExcelRangeBase cell, string Text, string author)
	{
		XmlElement xmlElement = CommentXml.CreateElement("comment", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		int row = cell.Start.Row;
		int col = cell.Start.Column;
		ExcelComment excelComment = null;
		if (Worksheet._commentsStore.NextCell(ref row, ref col))
		{
			excelComment = _list[Worksheet._commentsStore.GetValue(row, col)];
		}
		if (excelComment == null)
		{
			CommentXml.SelectSingleNode("d:comments/d:commentList", NameSpaceManager).AppendChild(xmlElement);
		}
		else
		{
			excelComment._commentHelper.TopNode.ParentNode.InsertBefore(xmlElement, excelComment._commentHelper.TopNode);
		}
		xmlElement.SetAttribute("ref", cell.Start.Address);
		ExcelComment excelComment2 = new ExcelComment(NameSpaceManager, xmlElement, cell);
		excelComment2.RichText.Add(Text);
		if (author != "")
		{
			excelComment2.Author = author;
		}
		_list.Add(excelComment2);
		Worksheet._commentsStore.SetValue(cell.Start.Row, cell.Start.Column, _list.Count - 1);
		if (!Worksheet.ExistsValueInner(cell._fromRow, cell._fromCol))
		{
			Worksheet.SetValueInner(cell._fromRow, cell._fromCol, null);
		}
		return excelComment2;
	}

	public void Remove(ExcelComment comment)
	{
		ulong cellID = ExcelCellBase.GetCellID(Worksheet.SheetID, comment.Range._fromRow, comment.Range._fromCol);
		int value = -1;
		ExcelComment excelComment = null;
		if (Worksheet._commentsStore.Exists(comment.Range._fromRow, comment.Range._fromCol, ref value))
		{
			excelComment = _list[value];
		}
		if (comment == excelComment)
		{
			comment.TopNode.ParentNode.RemoveChild(comment.TopNode);
			comment._commentHelper.TopNode.ParentNode.RemoveChild(comment._commentHelper.TopNode);
			Worksheet.VmlDrawingsComments._drawings.Delete(cellID);
			_list.RemoveAt(value);
			Worksheet._commentsStore.Delete(comment.Range._fromRow, comment.Range._fromCol, 1, 1, shift: false);
			CellsStoreEnumerator<int> cellsStoreEnumerator = new CellsStoreEnumerator<int>(Worksheet._commentsStore);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Value > value)
				{
					cellsStoreEnumerator.Value--;
				}
			}
			return;
		}
		throw new ArgumentException("Comment does not exist in the worksheet");
	}

	internal void Delete(int fromRow, int fromCol, int rows, int columns)
	{
		List<ExcelComment> list = new List<ExcelComment>();
		ExcelAddressBase excelAddressBase = null;
		foreach (ExcelComment item in _list)
		{
			excelAddressBase = new ExcelAddressBase(item.Address);
			if (fromCol > 0 && excelAddressBase._fromCol >= fromCol)
			{
				excelAddressBase = excelAddressBase.DeleteColumn(fromCol, columns);
			}
			if (fromRow > 0 && excelAddressBase._fromRow >= fromRow)
			{
				excelAddressBase = excelAddressBase.DeleteRow(fromRow, rows);
			}
			if (excelAddressBase == null || excelAddressBase.Address == "#REF!")
			{
				list.Add(item);
			}
			else
			{
				item.Reference = excelAddressBase.Address;
			}
		}
		foreach (ExcelComment item2 in list)
		{
			Remove(item2);
		}
	}

	public void Insert(int fromRow, int fromCol, int rows, int columns)
	{
		foreach (ExcelComment item in _list)
		{
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(item.Address);
			if (rows > 0 && excelAddressBase._fromRow >= fromRow)
			{
				item.Reference = item.Range.AddRow(fromRow, rows).Address;
			}
			if (columns > 0 && excelAddressBase._fromCol >= fromCol)
			{
				item.Reference = item.Range.AddColumn(fromCol, columns).Address;
			}
		}
	}

	void IDisposable.Dispose()
	{
	}

	public void RemoveAt(int Index)
	{
		Remove(this[Index]);
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	internal void Clear()
	{
		while (Count > 0)
		{
			RemoveAt(0);
		}
	}
}
