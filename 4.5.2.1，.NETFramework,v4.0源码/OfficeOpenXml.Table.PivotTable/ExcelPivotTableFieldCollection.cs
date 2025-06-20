using System;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
{
	public ExcelPivotTableField this[string name]
	{
		get
		{
			foreach (ExcelPivotTableField item in _list)
			{
				if (item.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return null;
		}
	}

	internal ExcelPivotTableFieldCollection(ExcelPivotTable table, string topNode)
		: base(table)
	{
	}

	public ExcelPivotTableField GetDateGroupField(eDateGroupBy GroupBy)
	{
		foreach (ExcelPivotTableField item in _list)
		{
			if (item.Grouping is ExcelPivotTableFieldDateGroup && ((ExcelPivotTableFieldDateGroup)item.Grouping).GroupBy == GroupBy)
			{
				return item;
			}
		}
		return null;
	}

	public ExcelPivotTableField GetNumericGroupField()
	{
		foreach (ExcelPivotTableField item in _list)
		{
			if (item.Grouping is ExcelPivotTableFieldNumericGroup)
			{
				return item;
			}
		}
		return null;
	}
}
