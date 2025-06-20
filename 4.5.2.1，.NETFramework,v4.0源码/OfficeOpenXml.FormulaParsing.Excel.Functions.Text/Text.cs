using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Text : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		object valueFirst = arguments.First().ValueFirst;
		string text = ArgToString(arguments, 1);
		text = text.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".");
		text = text.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator.Replace('\u00a0', ' '), ",");
		string format = context.ExcelDataProvider.GetFormat(valueFirst, text);
		return CreateResult(format, DataType.String);
	}
}
