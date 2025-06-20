using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Lookup : LookupFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		if (HaveTwoRanges(arguments))
		{
			return HandleTwoRanges(arguments, context);
		}
		return HandleSingleRange(arguments, context);
	}

	private bool HaveTwoRanges(IEnumerable<FunctionArgument> arguments)
	{
		if (arguments.Count() == 2)
		{
			return false;
		}
		return ExcelAddressUtil.IsValidAddress(arguments.ElementAt(2).Value.ToString());
	}

	private CompileResult HandleSingleRange(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		object value = arguments.ElementAt(0).Value;
		Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
		string text = ArgToString(arguments, 1);
		RangeAddress rangeAddress = new RangeAddressFactory(context.ExcelDataProvider).Create(text);
		int num = rangeAddress.ToRow - rangeAddress.FromRow;
		int num2 = rangeAddress.ToCol - rangeAddress.FromCol;
		int lookupIndex = num2 + 1;
		LookupDirection direction = LookupDirection.Vertical;
		if (num2 > num)
		{
			lookupIndex = num + 1;
			direction = LookupDirection.Horizontal;
		}
		LookupArguments lookupArguments = new LookupArguments(value, text, lookupIndex, 0, rangeLookup: true, arguments.ElementAt(1).ValueAsRangeInfo);
		LookupNavigator navigator = LookupNavigatorFactory.Create(direction, lookupArguments, context);
		return Lookup(navigator, lookupArguments);
	}

	private CompileResult HandleTwoRanges(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		object value = arguments.ElementAt(0).Value;
		Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
		Require.That(arguments.ElementAt(2).Value).Named("secondAddress").IsNotNull();
		string text = ArgToString(arguments, 1);
		string range = ArgToString(arguments, 2);
		RangeAddressFactory rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
		RangeAddress rangeAddress = rangeAddressFactory.Create(text);
		RangeAddress rangeAddress2 = rangeAddressFactory.Create(range);
		int lookupIndex = rangeAddress2.FromCol - rangeAddress.FromCol + 1;
		int lookupOffset = rangeAddress2.FromRow - rangeAddress.FromRow;
		LookupDirection lookupDirection = GetLookupDirection(rangeAddress);
		if (lookupDirection == LookupDirection.Horizontal)
		{
			lookupIndex = rangeAddress2.FromRow - rangeAddress.FromRow + 1;
			lookupOffset = rangeAddress2.FromCol - rangeAddress.FromCol;
		}
		LookupArguments lookupArguments = new LookupArguments(value, text, lookupIndex, lookupOffset, rangeLookup: true, arguments.ElementAt(1).ValueAsRangeInfo);
		LookupNavigator navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArguments, context);
		return Lookup(navigator, lookupArguments);
	}
}
