using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaList : ExcelDataValidationFormula, IExcelDataValidationFormulaList, IExcelDataValidationFormula
{
	private class DataValidationList : IList<string>, ICollection<string>, IEnumerable<string>, IEnumerable, ICollection
	{
		private IList<string> _items = new List<string>();

		private EventHandler<EventArgs> _listChanged;

		string IList<string>.this[int index]
		{
			get
			{
				return _items[index];
			}
			set
			{
				_items[index] = value;
				OnListChanged();
			}
		}

		int ICollection<string>.Count => _items.Count;

		bool ICollection<string>.IsReadOnly => false;

		int ICollection.Count => _items.Count;

		public bool IsSynchronized => ((ICollection)_items).IsSynchronized;

		public object SyncRoot => ((ICollection)_items).SyncRoot;

		public event EventHandler<EventArgs> ListChanged
		{
			add
			{
				_listChanged = (EventHandler<EventArgs>)Delegate.Combine(_listChanged, value);
			}
			remove
			{
				_listChanged = (EventHandler<EventArgs>)Delegate.Remove(_listChanged, value);
			}
		}

		private void OnListChanged()
		{
			if (_listChanged != null)
			{
				_listChanged(this, EventArgs.Empty);
			}
		}

		int IList<string>.IndexOf(string item)
		{
			return _items.IndexOf(item);
		}

		void IList<string>.Insert(int index, string item)
		{
			_items.Insert(index, item);
			OnListChanged();
		}

		void IList<string>.RemoveAt(int index)
		{
			_items.RemoveAt(index);
			OnListChanged();
		}

		void ICollection<string>.Add(string item)
		{
			_items.Add(item);
			OnListChanged();
		}

		void ICollection<string>.Clear()
		{
			_items.Clear();
			OnListChanged();
		}

		bool ICollection<string>.Contains(string item)
		{
			return _items.Contains(item);
		}

		void ICollection<string>.CopyTo(string[] array, int arrayIndex)
		{
			_items.CopyTo(array, arrayIndex);
		}

		bool ICollection<string>.Remove(string item)
		{
			bool result = _items.Remove(item);
			OnListChanged();
			return result;
		}

		IEnumerator<string> IEnumerable<string>.GetEnumerator()
		{
			return _items.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return _items.GetEnumerator();
		}

		public void CopyTo(Array array, int index)
		{
			_items.CopyTo((string[])array, index);
		}
	}

	private string _formulaPath;

	public IList<string> Values { get; private set; }

	public ExcelDataValidationFormulaList(XmlNamespaceManager namespaceManager, XmlNode itemNode, string formulaPath)
		: base(namespaceManager, itemNode, formulaPath)
	{
		Require.Argument(formulaPath).IsNotNullOrEmpty("formulaPath");
		_formulaPath = formulaPath;
		DataValidationList dataValidationList = new DataValidationList();
		dataValidationList.ListChanged += values_ListChanged;
		Values = dataValidationList;
		SetInitialValues();
	}

	private void SetInitialValues()
	{
		string xmlNodeString = GetXmlNodeString(_formulaPath);
		if (string.IsNullOrEmpty(xmlNodeString))
		{
			return;
		}
		if (xmlNodeString.StartsWith("\"") && xmlNodeString.EndsWith("\""))
		{
			xmlNodeString = xmlNodeString.TrimStart('"').TrimEnd('"');
			string[] array = xmlNodeString.Split(new char[1] { ',' }, StringSplitOptions.RemoveEmptyEntries);
			foreach (string item in array)
			{
				Values.Add(item);
			}
		}
		else
		{
			base.ExcelFormula = xmlNodeString;
		}
	}

	private void values_ListChanged(object sender, EventArgs e)
	{
		if (Values.Count > 0)
		{
			base.State = FormulaState.Value;
		}
		string valueAsString = GetValueAsString();
		if (valueAsString.Length > 255)
		{
			throw new InvalidOperationException("The total length of a DataValidation list cannot exceed 255 characters");
		}
		SetXmlNodeString(_formulaPath, valueAsString);
	}

	protected override string GetValueAsString()
	{
		StringBuilder stringBuilder = new StringBuilder();
		foreach (string value in Values)
		{
			if (stringBuilder.Length == 0)
			{
				stringBuilder.Append("\"");
				stringBuilder.Append(value);
			}
			else
			{
				stringBuilder.AppendFormat(",{0}", value);
			}
		}
		stringBuilder.Append("\"");
		return stringBuilder.ToString();
	}

	internal override void ResetValue()
	{
		Values.Clear();
	}
}
