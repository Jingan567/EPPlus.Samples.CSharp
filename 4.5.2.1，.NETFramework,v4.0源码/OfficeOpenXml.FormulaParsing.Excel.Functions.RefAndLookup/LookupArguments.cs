using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class LookupArguments
{
	public enum LookupArgumentDataType
	{
		ExcelRange,
		DataArray
	}

	private readonly ArgumentParsers _argumentParsers;

	public object SearchedValue { get; private set; }

	public string RangeAddress { get; private set; }

	public int LookupIndex { get; private set; }

	public int LookupOffset { get; private set; }

	public bool RangeLookup { get; private set; }

	public IEnumerable<FunctionArgument> DataArray { get; private set; }

	public ExcelDataProvider.IRangeInfo RangeInfo { get; private set; }

	public LookupArgumentDataType ArgumentDataType { get; private set; }

	public LookupArguments(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		: this(arguments, new ArgumentParsers(), context)
	{
	}

	public LookupArguments(IEnumerable<FunctionArgument> arguments, ArgumentParsers argumentParsers, ParsingContext context)
	{
		_argumentParsers = argumentParsers;
		SearchedValue = arguments.ElementAt(0).Value;
		object value = arguments.ElementAt(1).Value;
		if (value is IEnumerable<FunctionArgument> dataArray)
		{
			DataArray = dataArray;
			ArgumentDataType = LookupArgumentDataType.DataArray;
		}
		else if (value is ExcelDataProvider.IRangeInfo rangeInfo)
		{
			RangeAddress = (string.IsNullOrEmpty(rangeInfo.Address.WorkSheet) ? rangeInfo.Address.Address : ("'" + rangeInfo.Address.WorkSheet + "'!" + rangeInfo.Address.Address));
			RangeInfo = rangeInfo;
			ArgumentDataType = LookupArgumentDataType.ExcelRange;
		}
		else
		{
			RangeAddress = value.ToString();
			ArgumentDataType = LookupArgumentDataType.ExcelRange;
		}
		FunctionArgument functionArgument = arguments.ElementAt(2);
		if (functionArgument.DataType == DataType.ExcelAddress)
		{
			ExcelAddress excelAddress = new ExcelAddress(functionArgument.Value.ToString());
			object rangeValue = context.ExcelDataProvider.GetRangeValue(excelAddress.WorkSheet, excelAddress._fromRow, excelAddress._fromCol);
			LookupIndex = (int)_argumentParsers.GetParser(DataType.Integer).Parse(rangeValue);
		}
		else
		{
			LookupIndex = (int)_argumentParsers.GetParser(DataType.Integer).Parse(arguments.ElementAt(2).Value);
		}
		if (arguments.Count() > 3)
		{
			RangeLookup = (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(arguments.ElementAt(3).Value);
		}
		else
		{
			RangeLookup = true;
		}
	}

	public LookupArguments(object searchedValue, string rangeAddress, int lookupIndex, int lookupOffset, bool rangeLookup, ExcelDataProvider.IRangeInfo rangeInfo)
	{
		SearchedValue = searchedValue;
		RangeAddress = rangeAddress;
		RangeInfo = rangeInfo;
		LookupIndex = lookupIndex;
		LookupOffset = lookupOffset;
		RangeLookup = rangeLookup;
	}
}
