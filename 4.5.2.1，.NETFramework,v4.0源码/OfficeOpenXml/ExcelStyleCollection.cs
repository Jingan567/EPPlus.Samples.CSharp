using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public class ExcelStyleCollection<T> : IEnumerable<T>, IEnumerable
{
	private bool _setNextIdManual;

	internal List<T> _list = new List<T>();

	private Dictionary<string, int> _dic = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

	internal int NextId;

	public XmlNode TopNode { get; set; }

	public T this[int PositionID] => _list[PositionID];

	public int Count => _list.Count;

	public ExcelStyleCollection()
	{
		_setNextIdManual = false;
	}

	public ExcelStyleCollection(bool SetNextIdManual)
	{
		_setNextIdManual = SetNextIdManual;
	}

	public IEnumerator<T> GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	internal int Add(string key, T item)
	{
		_list.Add(item);
		if (!_dic.ContainsKey(key.ToLower(CultureInfo.InvariantCulture)))
		{
			_dic.Add(key.ToLower(CultureInfo.InvariantCulture), _list.Count - 1);
		}
		if (_setNextIdManual)
		{
			NextId++;
		}
		return _list.Count - 1;
	}

	internal bool FindByID(string key, ref T obj)
	{
		if (_dic.ContainsKey(key))
		{
			obj = _list[_dic[key.ToLower(CultureInfo.InvariantCulture)]];
			return true;
		}
		return false;
	}

	internal int FindIndexByID(string key)
	{
		if (_dic.ContainsKey(key))
		{
			return _dic[key];
		}
		return int.MinValue;
	}

	internal bool ExistsKey(string key)
	{
		return _dic.ContainsKey(key);
	}

	internal void Sort(Comparison<T> c)
	{
		_list.Sort(c);
	}
}
