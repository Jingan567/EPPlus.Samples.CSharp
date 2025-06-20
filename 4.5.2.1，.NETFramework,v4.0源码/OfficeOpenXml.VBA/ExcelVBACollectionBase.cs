using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml.VBA;

public class ExcelVBACollectionBase<T> : IEnumerable<T>, IEnumerable
{
	protected internal List<T> _list = new List<T>();

	public T this[string Name] => _list.Find((T f) => TypeCompat.GetPropertyValue(f, "Name").ToString().Equals(Name, StringComparison.OrdinalIgnoreCase));

	public T this[int Index] => _list[Index];

	public int Count => _list.Count;

	public IEnumerator<T> GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	public bool Exists(string Name)
	{
		return _list.Exists((T f) => TypeCompat.GetPropertyValue(f, "Name").ToString().Equals(Name, StringComparison.OrdinalIgnoreCase));
	}

	public void Remove(T Item)
	{
		_list.Remove(Item);
	}

	public void RemoveAt(int index)
	{
		_list.RemoveAt(index);
	}

	internal void Clear()
	{
		_list.Clear();
	}
}
