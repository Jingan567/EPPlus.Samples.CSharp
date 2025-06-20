using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Match : LookupFunction
{
	private enum MatchType
	{
		ClosestAbove = -1,
		ExactMatch,
		ClosestBelow
	}

	public Match()
		: base(new WildCardValueMatcher(), new CompileResultFactory())
	{
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		object value = arguments.ElementAt(0).Value;
		string text = ArgToAddress(arguments, 1);
		RangeAddress rangeAddress = new RangeAddressFactory(context.ExcelDataProvider).Create(text);
		MatchType matchType = GetMatchType(arguments);
		LookupArguments args = new LookupArguments(value, text, 0, 0, rangeLookup: false, arguments.ElementAt(1).ValueAsRangeInfo);
		LookupNavigator lookupNavigator = LookupNavigatorFactory.Create(GetLookupDirection(rangeAddress), args, context);
		int? num = null;
		do
		{
			int num2 = IsMatch(lookupNavigator.CurrentValue, value);
			if (num2 == 0)
			{
				return CreateResult(lookupNavigator.Index + 1, DataType.Integer);
			}
			if ((matchType == MatchType.ClosestBelow && num2 < 0) || (matchType == MatchType.ClosestAbove && num2 > 0))
			{
				num = lookupNavigator.Index + 1;
			}
			else if (matchType == MatchType.ClosestBelow || matchType == MatchType.ClosestAbove)
			{
				break;
			}
		}
		while (lookupNavigator.MoveNext());
		return CreateResult(num, DataType.Integer);
	}

	private MatchType GetMatchType(IEnumerable<FunctionArgument> arguments)
	{
		MatchType result = MatchType.ClosestBelow;
		if (arguments.Count() > 2)
		{
			result = (MatchType)ArgToInt(arguments, 2);
		}
		return result;
	}
}
