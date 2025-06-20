using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class SumIfs : MultipleRangeCriteriasFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 3);
		new List<int>();
		ExcelDataProvider.IRangeInfo valueAsRangeInfo = array[0].ValueAsRangeInfo;
		List<double> sumRange;
		if (valueAsRangeInfo != null)
		{
			sumRange = ArgsToDoubleEnumerableZeroPadded(ignoreHiddenCells: true, valueAsRangeInfo, context).ToList();
		}
		else
		{
			sumRange = ArgsToDoubleEnumerable(ignoreHiddenCells: true, new List<FunctionArgument> { array[0] }, context).Select((Func<ExcelDoubleCellValue, double>)((ExcelDoubleCellValue x) => x)).ToList();
		}
		List<ExcelDataProvider.IRangeInfo> list = new List<ExcelDataProvider.IRangeInfo>();
		List<string> list2 = new List<string>();
		for (int i = 1; i < 31 && array.Length > i; i += 2)
		{
			ExcelDataProvider.IRangeInfo valueAsRangeInfo2 = array[i].ValueAsRangeInfo;
			list.Add(valueAsRangeInfo2);
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
		double num = enumerable.Sum((int index) => sumRange[index]);
		return CreateResult(num, DataType.Decimal);
	}
}
