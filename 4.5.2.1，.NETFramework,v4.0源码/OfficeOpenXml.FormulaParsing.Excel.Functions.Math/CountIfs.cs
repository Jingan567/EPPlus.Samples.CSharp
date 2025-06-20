using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class CountIfs : MultipleRangeCriteriasFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		List<ExcelDataProvider.IRangeInfo> list = new List<ExcelDataProvider.IRangeInfo>();
		List<string> list2 = new List<string>();
		for (int i = 0; i < 30 && array.Length > i; i += 2)
		{
			ExcelDataProvider.IRangeInfo valueAsRangeInfo = array[i].ValueAsRangeInfo;
			list.Add(valueAsRangeInfo);
			string item = ((array[i + 1].Value != null) ? array[i + 1].Value.ToString() : null);
			list2.Add(item);
		}
		IEnumerable<int> enumerable = GetMatchIndexes(list[0], list2[0]);
		IList<int> list3 = (enumerable as IList<int>) ?? enumerable.ToList();
		for (int j = 1; j < list.Count; j++)
		{
			if (!list3.Any())
			{
				break;
			}
			List<int> matchIndexes = GetMatchIndexes(list[j], list2[j]);
			enumerable = list3.Intersect(matchIndexes);
		}
		return CreateResult((double)enumerable.Count(), DataType.Integer);
	}
}
