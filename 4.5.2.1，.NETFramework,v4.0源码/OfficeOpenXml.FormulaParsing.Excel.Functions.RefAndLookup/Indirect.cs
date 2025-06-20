using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Indirect : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		string address = ArgToAddress(arguments, 0);
		string text = new ExcelAddress(address).WorkSheet;
		if (string.IsNullOrEmpty(text))
		{
			text = context.Scopes.Current.Address.Worksheet;
		}
		ExcelDataProvider.IRangeInfo range = context.ExcelDataProvider.GetRange(text, address);
		if (range.IsEmpty)
		{
			return CompileResult.Empty;
		}
		return new CompileResult(range, DataType.Enumerable);
	}
}
