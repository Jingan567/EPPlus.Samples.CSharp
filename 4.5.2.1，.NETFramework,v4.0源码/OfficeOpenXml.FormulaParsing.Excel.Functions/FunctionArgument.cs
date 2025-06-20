using System;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class FunctionArgument
{
	private ExcelCellState _excelCellState;

	public object Value { get; private set; }

	public DataType DataType { get; }

	public Type Type
	{
		get
		{
			if (Value == null)
			{
				return null;
			}
			return Value.GetType();
		}
	}

	public bool IsExcelRange
	{
		get
		{
			if (Value != null)
			{
				return Value is ExcelDataProvider.IRangeInfo;
			}
			return false;
		}
	}

	public bool ValueIsExcelError => ExcelErrorValue.Values.IsErrorValue(Value);

	public ExcelErrorValue ValueAsExcelErrorValue => ExcelErrorValue.Parse(Value.ToString());

	public ExcelDataProvider.IRangeInfo ValueAsRangeInfo => Value as ExcelDataProvider.IRangeInfo;

	public object ValueFirst
	{
		get
		{
			if (Value is ExcelDataProvider.INameInfo)
			{
				Value = ((ExcelDataProvider.INameInfo)Value).Value;
			}
			if (!(Value is ExcelDataProvider.IRangeInfo rangeInfo))
			{
				return Value;
			}
			return rangeInfo.GetValue(rangeInfo.Address._fromRow, rangeInfo.Address._fromCol);
		}
	}

	public FunctionArgument(object val)
	{
		Value = val;
		DataType = DataType.Unknown;
	}

	public FunctionArgument(object val, DataType dataType)
		: this(val)
	{
		DataType = dataType;
	}

	public void SetExcelStateFlag(ExcelCellState state)
	{
		_excelCellState |= state;
	}

	public bool ExcelStateFlagIsSet(ExcelCellState state)
	{
		return (_excelCellState & state) != 0;
	}
}
