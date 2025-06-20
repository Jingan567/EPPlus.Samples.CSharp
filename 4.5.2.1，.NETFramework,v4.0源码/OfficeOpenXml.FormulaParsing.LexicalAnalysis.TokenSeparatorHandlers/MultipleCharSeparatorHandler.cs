namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;

public class MultipleCharSeparatorHandler : SeparatorHandler
{
	private ITokenSeparatorProvider _tokenSeparatorProvider;

	public MultipleCharSeparatorHandler()
		: this(new TokenSeparatorProvider())
	{
	}

	public MultipleCharSeparatorHandler(ITokenSeparatorProvider tokenSeparatorProvider)
	{
		_tokenSeparatorProvider = tokenSeparatorProvider;
	}

	public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
	{
		if (IsPartOfMultipleCharSeparator(context, c))
		{
			string key = context.LastToken.Value + c;
			Token newToken = _tokenSeparatorProvider.Tokens[key];
			context.ReplaceLastToken(newToken);
			context.NewToken();
			return true;
		}
		return false;
	}

	private bool IsPartOfMultipleCharSeparator(TokenizerContext context, char c)
	{
		string item = ((context.LastToken != null) ? context.LastToken.Value : string.Empty);
		if (_tokenSeparatorProvider.IsOperator(item) && _tokenSeparatorProvider.IsPossibleLastPartOfMultipleCharOperator(c.ToString()))
		{
			return !context.CurrentTokenHasValue;
		}
		return false;
	}
}
