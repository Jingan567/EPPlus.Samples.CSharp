using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Time : TimeBaseFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		string input = arguments.ElementAt(0).Value.ToString();
		if (arguments.Count() == 1 && base.TimeStringParser.CanParse(input))
		{
			return new CompileResult(base.TimeStringParser.Parse(input), DataType.Time);
		}
		ValidateArguments(arguments, 3);
		int hour = ArgToInt(arguments, 0);
		int min = ArgToInt(arguments, 1);
		int sec = ArgToInt(arguments, 2);
		ThrowArgumentExceptionIf(() => sec < 0 || sec > 59, "Invalid second: " + sec);
		ThrowArgumentExceptionIf(() => min < 0 || min > 59, "Invalid minute: " + min);
		ThrowArgumentExceptionIf(() => min < 0 || hour > 23, "Invalid hour: " + hour);
		double seconds = hour * 60 * 60 + min * 60 + sec;
		return CreateResult(GetTimeSerialNumber(seconds), DataType.Time);
	}
}
