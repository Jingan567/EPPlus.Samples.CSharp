using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Minute : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object value = arguments.ElementAt(0).Value;
		System.DateTime minValue = System.DateTime.MinValue;
		return CreateResult(((!(value is string)) ? System.DateTime.FromOADate(ArgToDecimal(arguments, 0)) : System.DateTime.Parse(value.ToString())).Minute, DataType.Integer);
	}
}
