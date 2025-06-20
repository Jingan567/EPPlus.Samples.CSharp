namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;

public static class TokenSeparatorHandler
{
	private static SeparatorHandler[] _handlers = new SeparatorHandler[4]
	{
		new StringHandler(),
		new BracketHandler(),
		new SheetnameHandler(),
		new MultipleCharSeparatorHandler()
	};

	public static bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
	{
		SeparatorHandler[] handlers = _handlers;
		for (int i = 0; i < handlers.Length; i++)
		{
			if (handlers[i].Handle(c, tokenSeparator, context, tokenIndexProvider))
			{
				return true;
			}
		}
		return false;
	}
}
