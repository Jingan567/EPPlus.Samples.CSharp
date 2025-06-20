using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class CellReferenceProvider
{
	public virtual IEnumerable<string> GetReferencedAddresses(string cellFormula, ParsingContext context)
	{
		List<string> list = new List<string>();
		foreach (Token item in from x in context.Configuration.Lexer.Tokenize(cellFormula, context.Scopes.Current.Address.Worksheet)
			where x.TokenType == TokenType.ExcelAddress
			select x)
		{
			RangeAddress rangeAddress = context.RangeAddressFactory.Create(item.Value);
			List<string> list2 = new List<string>();
			if (rangeAddress.FromRow < rangeAddress.ToRow || rangeAddress.FromCol < rangeAddress.ToCol)
			{
				for (int i = rangeAddress.FromCol; i <= rangeAddress.ToCol; i++)
				{
					for (int j = rangeAddress.FromRow; j <= rangeAddress.ToRow; j++)
					{
						list.Add(context.RangeAddressFactory.Create(i, j).Address);
					}
				}
			}
			else
			{
				list2.Add(item.Value);
			}
			list.AddRange(list2);
		}
		return list;
	}
}
