namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;

public abstract class SeparatorHandler
{
	protected bool IsDoubleQuote(Token tokenSeparator, int formulaCharIndex, TokenizerContext context)
	{
		if (tokenSeparator.TokenType == TokenType.String && formulaCharIndex + 1 < context.FormulaChars.Length)
		{
			return context.FormulaChars[formulaCharIndex + 1] == '"';
		}
		return false;
	}

	public abstract bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider);
}
