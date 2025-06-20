using System;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class CompileResult
{
	private static CompileResult _empty = new CompileResult(null, DataType.Empty);

	private double? _ResultNumeric;

	public static CompileResult Empty => _empty;

	public object Result { get; private set; }

	public object ResultValue
	{
		get
		{
			if (!(Result is ExcelDataProvider.IRangeInfo rangeInfo))
			{
				return Result;
			}
			return rangeInfo.GetValue(rangeInfo.Address._fromRow, rangeInfo.Address._fromCol);
		}
	}

	public double ResultNumeric
	{
		get
		{
			if (!_ResultNumeric.HasValue)
			{
				if (IsNumeric)
				{
					_ResultNumeric = ((Result == null) ? 0.0 : Convert.ToDouble(Result));
				}
				else if (Result is DateTime)
				{
					_ResultNumeric = ((DateTime)Result).ToOADate();
				}
				else if (Result is TimeSpan)
				{
					_ResultNumeric = DateTime.FromOADate(0.0).Add((TimeSpan)Result).ToOADate();
				}
				else
				{
					if (Result is ExcelDataProvider.IRangeInfo)
					{
						return ((ExcelDataProvider.IRangeInfo)Result).FirstOrDefault()?.ValueDoubleLogical ?? 0.0;
					}
					if (!IsNumericString && !IsDateString)
					{
						_ResultNumeric = 0.0;
					}
				}
			}
			return _ResultNumeric.Value;
		}
	}

	public DataType DataType { get; private set; }

	public bool IsNumeric
	{
		get
		{
			if (DataType != DataType.Decimal && DataType != 0 && DataType != DataType.Empty && DataType != DataType.Boolean)
			{
				return DataType == DataType.Date;
			}
			return true;
		}
	}

	public bool IsNumericString
	{
		get
		{
			if (DataType == DataType.String && ConvertUtil.TryParseNumericString(Result, out var result))
			{
				_ResultNumeric = result;
				return true;
			}
			return false;
		}
	}

	public bool IsDateString
	{
		get
		{
			if (DataType == DataType.String && ConvertUtil.TryParseDateString(Result, out var result))
			{
				_ResultNumeric = result.ToOADate();
				return true;
			}
			return false;
		}
	}

	public bool IsResultOfSubtotal { get; set; }

	public bool IsHiddenCell { get; set; }

	public CompileResult(object result, DataType dataType)
	{
		if (result is ExcelDoubleCellValue)
		{
			Result = ((ExcelDoubleCellValue)result).Value;
		}
		else
		{
			Result = result;
		}
		DataType = dataType;
	}

	public CompileResult(eErrorType errorType)
	{
		Result = ExcelErrorValue.Create(errorType);
		DataType = DataType.ExcelError;
	}

	public CompileResult(ExcelErrorValue errorValue)
	{
		Require.Argument(errorValue).IsNotNull("errorValue");
		Result = errorValue;
		DataType = DataType.ExcelError;
	}
}
