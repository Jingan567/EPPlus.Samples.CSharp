using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Proper : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		string text = ArgToString(arguments, 0).ToLower(CultureInfo.InvariantCulture);
		StringBuilder stringBuilder = new StringBuilder();
		char c = '.';
		string text2 = text;
		for (int i = 0; i < text2.Length; i++)
		{
			char c2 = text2[i];
			if (!char.IsLetter(c))
			{
				stringBuilder.Append(ConvertUtil._invariantTextInfo.ToUpper(c2.ToString()));
			}
			else
			{
				stringBuilder.Append(c2);
			}
			c = c2;
		}
		return CreateResult(stringBuilder.ToString(), DataType.String);
	}
}
