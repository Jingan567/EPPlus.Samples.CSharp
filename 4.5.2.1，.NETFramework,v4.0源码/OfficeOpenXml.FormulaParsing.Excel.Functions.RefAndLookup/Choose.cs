using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Choose : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		List<object> items = new List<object>();
		for (int i = 0; i < arguments.Count(); i++)
		{
			items.Add(arguments.ElementAt(i).ValueFirst);
		}
		if (arguments.ElementAt(0).ValueFirst is IEnumerable<FunctionArgument> source && source.Count() > 1)
		{
			IntArgumentParser intParser = new IntArgumentParser();
			object[] result = source.Select((FunctionArgument chosenIndex) => items[(int)intParser.Parse(chosenIndex.ValueFirst)]).ToArray();
			return CreateResult(result, DataType.Enumerable);
		}
		int index = ArgToInt(arguments, 0);
		return CreateResult(items[index].ToString(), DataType.String);
	}
}
