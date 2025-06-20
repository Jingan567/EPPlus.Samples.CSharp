using System;
using System.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

internal static class CellStateHelper
{
	private static bool IsSubTotal(ExcelDataProvider.ICellInfo c)
	{
		if (c.Tokens == null)
		{
			return false;
		}
		return c.Tokens.Any((Token token) => token.TokenType == TokenType.Function && token.Value.Equals("SUBTOTAL", StringComparison.OrdinalIgnoreCase));
	}

	internal static bool ShouldIgnore(bool ignoreHiddenValues, ExcelDataProvider.ICellInfo c, ParsingContext context)
	{
		if (!ignoreHiddenValues || !c.IsHiddenRow)
		{
			if (context.Scopes.Current.IsSubtotal)
			{
				return IsSubTotal(c);
			}
			return false;
		}
		return true;
	}

	internal static bool ShouldIgnore(bool ignoreHiddenValues, FunctionArgument arg, ParsingContext context)
	{
		if (ignoreHiddenValues)
		{
			return arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell);
		}
		return false;
	}
}
