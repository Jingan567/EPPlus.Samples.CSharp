using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class ErrorType : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		FunctionArgument functionArgument = arguments.ElementAt(0);
		if (!(bool)context.Configuration.FunctionRepository.GetFunction("iserror").Execute(arguments, context).Result)
		{
			return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
		}
		return functionArgument.ValueAsExcelErrorValue.Type switch
		{
			eErrorType.Null => CreateResult(1, DataType.Integer), 
			eErrorType.Div0 => CreateResult(2, DataType.Integer), 
			eErrorType.Value => CreateResult(3, DataType.Integer), 
			eErrorType.Ref => CreateResult(4, DataType.Integer), 
			eErrorType.Name => CreateResult(5, DataType.Integer), 
			eErrorType.Num => CreateResult(6, DataType.Integer), 
			eErrorType.NA => CreateResult(7, DataType.Integer), 
			_ => CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError), 
		};
	}
}
