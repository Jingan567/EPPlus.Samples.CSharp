using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Weekday : ExcelFunction
{
	private static List<int> _oneBasedStartOnSunday = new List<int> { 1, 2, 3, 4, 5, 6, 7 };

	private static List<int> _oneBasedStartOnMonday = new List<int> { 7, 1, 2, 3, 4, 5, 6 };

	private static List<int> _zeroBasedStartOnSunday = new List<int> { 6, 0, 1, 2, 3, 4, 5 };

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double d = ArgToDecimal(arguments, 0);
		int returnType = ((arguments.Count() <= 1) ? 1 : ArgToInt(arguments, 1));
		return CreateResult(CalculateDayOfWeek(System.DateTime.FromOADate(d), returnType), DataType.Integer);
	}

	private int CalculateDayOfWeek(System.DateTime dateTime, int returnType)
	{
		int dayOfWeek = (int)dateTime.DayOfWeek;
		return returnType switch
		{
			1 => _oneBasedStartOnSunday[dayOfWeek], 
			2 => _oneBasedStartOnMonday[dayOfWeek], 
			3 => _zeroBasedStartOnSunday[dayOfWeek], 
			_ => throw new ExcelErrorValueException(eErrorType.Num), 
		};
	}
}
