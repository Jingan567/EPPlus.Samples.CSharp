using System;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableRowColumnFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
{
	internal string _topNode;

	internal ExcelPivotTableRowColumnFieldCollection(ExcelPivotTable table, string topNode)
		: base(table)
	{
		_topNode = topNode;
	}

	public ExcelPivotTableField Add(ExcelPivotTableField Field)
	{
		SetFlag(Field, value: true);
		_list.Add(Field);
		return Field;
	}

	internal ExcelPivotTableField Insert(ExcelPivotTableField Field, int Index)
	{
		SetFlag(Field, value: true);
		_list.Insert(Index, Field);
		return Field;
	}

	private void SetFlag(ExcelPivotTableField field, bool value)
	{
		string topNode = _topNode;
		switch (topNode)
		{
		default:
			_ = topNode == "dataFields";
			break;
		case "rowFields":
			if (field.IsColumnField || field.IsPageField)
			{
				throw new Exception("This field is a column or page field. Can't add it to the RowFields collection");
			}
			field.IsRowField = value;
			field.Axis = ePivotFieldAxis.Row;
			break;
		case "colFields":
			if (field.IsRowField || field.IsPageField)
			{
				throw new Exception("This field is a row or page field. Can't add it to the ColumnFields collection");
			}
			field.IsColumnField = value;
			field.Axis = ePivotFieldAxis.Column;
			break;
		case "pageFields":
			if (field.IsColumnField || field.IsRowField)
			{
				throw new Exception("Field is a column or row field. Can't add it to the PageFields collection");
			}
			if (_table.Address._fromRow < 3)
			{
				throw new Exception($"A pivot table with page fields must be located above row 3. Currenct location is {_table.Address.Address}");
			}
			field.IsPageField = value;
			field.Axis = ePivotFieldAxis.Page;
			break;
		}
	}

	public void Remove(ExcelPivotTableField Field)
	{
		if (!_list.Contains(Field))
		{
			throw new ArgumentException("Field not in collection");
		}
		SetFlag(Field, value: false);
		_list.Remove(Field);
	}

	public void RemoveAt(int Index)
	{
		if (Index > -1 && Index < _list.Count)
		{
			throw new IndexOutOfRangeException();
		}
		SetFlag(_list[Index], value: false);
		_list.RemoveAt(Index);
	}
}
