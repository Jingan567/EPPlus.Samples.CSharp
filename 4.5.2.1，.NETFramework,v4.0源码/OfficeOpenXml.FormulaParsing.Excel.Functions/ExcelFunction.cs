using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public abstract class ExcelFunction
{
	private readonly ArgumentCollectionUtil _argumentCollectionUtil;

	private readonly ArgumentParsers _argumentParsers;

	private readonly CompileResultValidators _compileResultValidators;

	public virtual bool IsLookupFuction => false;

	public virtual bool IsErrorHandlingFunction => false;

	public bool SkipArgumentEvaluation { get; set; }

	public ExcelFunction()
		: this(new ArgumentCollectionUtil(), new ArgumentParsers(), new CompileResultValidators())
	{
	}

	public ExcelFunction(ArgumentCollectionUtil argumentCollectionUtil, ArgumentParsers argumentParsers, CompileResultValidators compileResultValidators)
	{
		_argumentCollectionUtil = argumentCollectionUtil;
		_argumentParsers = argumentParsers;
		_compileResultValidators = compileResultValidators;
	}

	public abstract CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context);

	public virtual void BeforeInvoke(ParsingContext context)
	{
	}

	protected object GetFirstValue(IEnumerable<FunctionArgument> val)
	{
		FunctionArgument functionArgument = val.FirstOrDefault();
		if (functionArgument.Value is ExcelDataProvider.IRangeInfo)
		{
			ExcelDataProvider.IRangeInfo valueAsRangeInfo = functionArgument.ValueAsRangeInfo;
			return valueAsRangeInfo.GetValue(valueAsRangeInfo.Address._fromRow, valueAsRangeInfo.Address._fromCol);
		}
		return functionArgument?.Value;
	}

	protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength, eErrorType errorTypeToThrow)
	{
		Require.That(arguments).Named("arguments").IsNotNull();
		ThrowExcelErrorValueExceptionIf(delegate
		{
			int num = 0;
			if (arguments.Any())
			{
				foreach (FunctionArgument argument in arguments)
				{
					num++;
					if (num >= minLength)
					{
						return false;
					}
					if (argument.IsExcelRange)
					{
						num += argument.ValueAsRangeInfo.GetNCells();
						if (num >= minLength)
						{
							return false;
						}
					}
				}
			}
			return true;
		}, errorTypeToThrow);
	}

	protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength)
	{
		Require.That(arguments).Named("arguments").IsNotNull();
		ThrowArgumentExceptionIf(delegate
		{
			int num = 0;
			if (arguments.Any())
			{
				foreach (FunctionArgument argument in arguments)
				{
					num++;
					if (num >= minLength)
					{
						return false;
					}
					if (argument.IsExcelRange)
					{
						num += argument.ValueAsRangeInfo.GetNCells();
						if (num >= minLength)
						{
							return false;
						}
					}
				}
			}
			return true;
		}, "Expecting at least {0} arguments", minLength.ToString());
	}

	protected string ArgToAddress(IEnumerable<FunctionArgument> arguments, int index)
	{
		if (!arguments.ElementAt(index).IsExcelRange)
		{
			return ArgToString(arguments, index);
		}
		return arguments.ElementAt(index).ValueAsRangeInfo.Address.FullAddress;
	}

	protected int ArgToInt(IEnumerable<FunctionArgument> arguments, int index)
	{
		object valueFirst = arguments.ElementAt(index).ValueFirst;
		return (int)_argumentParsers.GetParser(DataType.Integer).Parse(valueFirst);
	}

	protected string ArgToString(IEnumerable<FunctionArgument> arguments, int index)
	{
		object valueFirst = arguments.ElementAt(index).ValueFirst;
		if (valueFirst == null)
		{
			return string.Empty;
		}
		return valueFirst.ToString();
	}

	protected double ArgToDecimal(object obj)
	{
		return (double)_argumentParsers.GetParser(DataType.Decimal).Parse(obj);
	}

	protected double ArgToDecimal(IEnumerable<FunctionArgument> arguments, int index)
	{
		return ArgToDecimal(arguments.ElementAt(index).Value);
	}

	protected double Divide(double left, double right)
	{
		if (System.Math.Abs(right - 0.0) < double.Epsilon)
		{
			throw new ExcelErrorValueException(eErrorType.Div0);
		}
		return left / right;
	}

	protected bool IsNumericString(object value)
	{
		if (value == null || string.IsNullOrEmpty(value.ToString()))
		{
			return false;
		}
		return Regex.IsMatch(value.ToString(), "^[\\d]+(\\,[\\d])?");
	}

	protected bool ArgToBool(IEnumerable<FunctionArgument> arguments, int index)
	{
		object obj = arguments.ElementAt(index).Value ?? string.Empty;
		return (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(obj);
	}

	protected void ThrowArgumentExceptionIf(Func<bool> condition, string message)
	{
		if (condition())
		{
			throw new ArgumentException(message);
		}
	}

	protected void ThrowArgumentExceptionIf(Func<bool> condition, string message, params object[] formats)
	{
		message = string.Format(message, formats);
		ThrowArgumentExceptionIf(condition, message);
	}

	protected void ThrowExcelErrorValueException(eErrorType errorType)
	{
		throw new ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(errorType));
	}

	protected void ThrowExcelErrorValueExceptionIf(Func<bool> condition, eErrorType errorType)
	{
		if (condition())
		{
			throw new ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(errorType));
		}
	}

	protected bool IsNumeric(object val)
	{
		if (val == null)
		{
			return false;
		}
		if (!TypeCompat.IsPrimitive(val) && !(val is double) && !(val is decimal) && !(val is System.DateTime))
		{
			return val is TimeSpan;
		}
		return true;
	}

	protected bool AreEqual(double d1, double d2)
	{
		return System.Math.Abs(d1 - d2) < double.Epsilon;
	}

	protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return ArgsToDoubleEnumerable(ignoreHiddenCells: false, arguments, context);
	}

	protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return _argumentCollectionUtil.ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors, arguments, context);
	}

	protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors: true, arguments, context);
	}

	protected virtual IEnumerable<double> ArgsToDoubleEnumerableZeroPadded(bool ignoreHiddenCells, ExcelDataProvider.IRangeInfo rangeInfo, ParsingContext context)
	{
		int row = rangeInfo.Address.Start.Row;
		int row2 = rangeInfo.Address.End.Row;
		FunctionArgument item = new FunctionArgument(rangeInfo);
		IEnumerable<ExcelDoubleCellValue> source = ArgsToDoubleEnumerable(ignoreHiddenCells, new List<FunctionArgument> { item }, context);
		Dictionary<int, double> dict = new Dictionary<int, double>();
		source.ToList().ForEach(delegate(ExcelDoubleCellValue x)
		{
			dict.Add(x.CellRow.Value, x.Value);
		});
		List<double> list = new List<double>();
		for (int i = row; i <= row2; i++)
		{
			if (dict.ContainsKey(i))
			{
				list.Add(dict[i]);
			}
			else
			{
				list.Add(0.0);
			}
		}
		return list;
	}

	protected virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return _argumentCollectionUtil.ArgsToObjectEnumerable(ignoreHiddenCells, arguments, context);
	}

	protected CompileResult CreateResult(object result, DataType dataType)
	{
		_compileResultValidators.GetValidator(dataType).Validate(result);
		return new CompileResult(result, dataType);
	}

	protected virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument, double, double> action)
	{
		return _argumentCollectionUtil.CalculateCollection(collection, result, action);
	}

	protected void CheckForAndHandleExcelError(FunctionArgument arg)
	{
		if (arg.ValueIsExcelError)
		{
			throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
		}
	}

	protected void CheckForAndHandleExcelError(ExcelDataProvider.ICellInfo cell)
	{
		if (cell.IsExcelError)
		{
			throw new ExcelErrorValueException(ExcelErrorValue.Parse(cell.Value.ToString()));
		}
	}

	protected CompileResult GetResultByObject(object result)
	{
		if (IsNumeric(result))
		{
			return CreateResult(result, DataType.Decimal);
		}
		if (result is string)
		{
			return CreateResult(result, DataType.String);
		}
		if (ExcelErrorValue.Values.IsErrorValue(result))
		{
			return CreateResult(result, DataType.ExcelAddress);
		}
		if (result == null)
		{
			return CompileResult.Empty;
		}
		return CreateResult(result, DataType.Enumerable);
	}
}
