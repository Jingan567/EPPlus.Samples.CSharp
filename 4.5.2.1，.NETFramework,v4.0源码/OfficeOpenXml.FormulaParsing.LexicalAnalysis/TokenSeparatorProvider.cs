using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class TokenSeparatorProvider : ITokenSeparatorProvider
{
	private static readonly Dictionary<string, Token> _tokens;

	IDictionary<string, Token> ITokenSeparatorProvider.Tokens => _tokens;

	static TokenSeparatorProvider()
	{
		_tokens = new Dictionary<string, Token>();
		_tokens.Add("+", new Token("+", TokenType.Operator));
		_tokens.Add("-", new Token("-", TokenType.Operator));
		_tokens.Add("*", new Token("*", TokenType.Operator));
		_tokens.Add("/", new Token("/", TokenType.Operator));
		_tokens.Add("^", new Token("^", TokenType.Operator));
		_tokens.Add("&", new Token("&", TokenType.Operator));
		_tokens.Add(">", new Token(">", TokenType.Operator));
		_tokens.Add("<", new Token("<", TokenType.Operator));
		_tokens.Add("=", new Token("=", TokenType.Operator));
		_tokens.Add("<=", new Token("<=", TokenType.Operator));
		_tokens.Add(">=", new Token(">=", TokenType.Operator));
		_tokens.Add("<>", new Token("<>", TokenType.Operator));
		_tokens.Add("(", new Token("(", TokenType.OpeningParenthesis));
		_tokens.Add(")", new Token(")", TokenType.ClosingParenthesis));
		_tokens.Add("{", new Token("{", TokenType.OpeningEnumerable));
		_tokens.Add("}", new Token("}", TokenType.ClosingEnumerable));
		_tokens.Add("'", new Token("'", TokenType.WorksheetName));
		_tokens.Add("\"", new Token("\"", TokenType.String));
		_tokens.Add(",", new Token(",", TokenType.Comma));
		_tokens.Add(";", new Token(";", TokenType.SemiColon));
		_tokens.Add("[", new Token("[", TokenType.OpeningBracket));
		_tokens.Add("]", new Token("]", TokenType.ClosingBracket));
		_tokens.Add("%", new Token("%", TokenType.Percent));
	}

	public bool IsOperator(string item)
	{
		if (_tokens.TryGetValue(item, out var value) && value.TokenType == TokenType.Operator)
		{
			return true;
		}
		return false;
	}

	public bool IsPossibleLastPartOfMultipleCharOperator(string part)
	{
		if (!(part == "="))
		{
			return part == ">";
		}
		return true;
	}
}
