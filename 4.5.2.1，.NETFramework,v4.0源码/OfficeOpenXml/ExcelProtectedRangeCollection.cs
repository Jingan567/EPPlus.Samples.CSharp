using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public class ExcelProtectedRangeCollection : XmlHelper, IEnumerable<ExcelProtectedRange>, IEnumerable
{
	private List<ExcelProtectedRange> _baseList = new List<ExcelProtectedRange>();

	public int Count => _baseList.Count;

	public ExcelProtectedRange this[int index] => _baseList[index];

	internal ExcelProtectedRangeCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
		: base(nsm, topNode)
	{
		base.SchemaNodeOrder = ws.SchemaNodeOrder;
		foreach (XmlNode item in topNode.SelectNodes("d:protectedRanges/d:protectedRange", nsm))
		{
			if (item is XmlElement)
			{
				_baseList.Add(new ExcelProtectedRange(item.Attributes["name"].Value, new ExcelAddress(SqRefUtility.FromSqRefAddress(item.Attributes["sqref"].Value)), nsm, topNode));
			}
		}
	}

	public ExcelProtectedRange Add(string name, ExcelAddress address)
	{
		if (!ExistNode("d:protectedRanges"))
		{
			CreateNode("d:protectedRanges");
		}
		foreach (ExcelProtectedRange @base in _baseList)
		{
			if (name.Equals(@base.Name, StringComparison.CurrentCultureIgnoreCase))
			{
				throw new InvalidOperationException($"A protected range with the namn {name} already exists");
			}
		}
		XmlElement xmlElement = base.TopNode.OwnerDocument.CreateElement("protectedRange", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		base.TopNode.SelectSingleNode("d:protectedRanges", base.NameSpaceManager).AppendChild(xmlElement);
		ExcelProtectedRange excelProtectedRange = new ExcelProtectedRange(name, address, base.NameSpaceManager, xmlElement);
		_baseList.Add(excelProtectedRange);
		return excelProtectedRange;
	}

	public void Clear()
	{
		DeleteNode("d:protectedRanges");
		_baseList.Clear();
	}

	public bool Contains(ExcelProtectedRange item)
	{
		return _baseList.Contains(item);
	}

	public void CopyTo(ExcelProtectedRange[] array, int arrayIndex)
	{
		_baseList.CopyTo(array, arrayIndex);
	}

	public bool Remove(ExcelProtectedRange item)
	{
		DeleteAllNode("d:protectedRanges/d:protectedRange[@name='" + item.Name + "' and @sqref='" + item.Address.Address + "']");
		if (_baseList.Count == 0)
		{
			DeleteNode("d:protectedRanges");
		}
		return _baseList.Remove(item);
	}

	public int IndexOf(ExcelProtectedRange item)
	{
		return _baseList.IndexOf(item);
	}

	public void RemoveAt(int index)
	{
		_baseList.RemoveAt(index);
	}

	IEnumerator<ExcelProtectedRange> IEnumerable<ExcelProtectedRange>.GetEnumerator()
	{
		return _baseList.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _baseList.GetEnumerator();
	}
}
