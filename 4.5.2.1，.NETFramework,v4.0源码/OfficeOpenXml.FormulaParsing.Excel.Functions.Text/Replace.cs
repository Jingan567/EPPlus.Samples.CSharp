using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Replace : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 4);
		string text = ArgToString(arguments, 0);
		int startPos = ArgToInt(arguments, 1);
		int nCharactersToReplace = ArgToInt(arguments, 2);
		string text2 = ArgToString(arguments, 3);
		string firstPart = GetFirstPart(text, startPos);
		string lastPart = GetLastPart(text, startPos, nCharactersToReplace);
		string result = firstPart + text2 + lastPart;
		return CreateResult(result, DataType.String);
	}

	private string GetFirstPart(string text, int startPos)
	{
		return text.Substring(0, startPos - 1);
	}

	private string GetLastPart(string text, int startPos, int nCharactersToReplace)
	{
		int num = startPos - 1;
		num += nCharactersToReplace;
		return text.Substring(num, text.Length - num);
	}
}
