using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class TokenHandler : ITokenIndexProvider
{
	private readonly TokenizerContext _context;

	private readonly ITokenSeparatorProvider _tokenProvider;

	private readonly ITokenFactory _tokenFactory;

	private int _tokenIndex = -1;

	public string Worksheet { get; set; }

	int ITokenIndexProvider.Index => _tokenIndex;

	public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
	{
		_context = context;
		_tokenFactory = tokenFactory;
		_tokenProvider = tokenProvider;
	}

	public bool HasMore()
	{
		return _tokenIndex < _context.FormulaChars.Length - 1;
	}

	public void Next()
	{
		_tokenIndex++;
		Handle();
	}

	private void Handle()
	{
		char c = _context.FormulaChars[_tokenIndex];
		if (CharIsTokenSeparator(c, out var token))
		{
			if (TokenSeparatorHandler.Handle(c, token, _context, this))
			{
				return;
			}
			if (_context.CurrentTokenHasValue)
			{
				if (Regex.IsMatch(_context.CurrentToken, "^\"*$"))
				{
					_context.AddToken(_tokenFactory.Create(_context.CurrentToken, TokenType.StringContent));
				}
				else
				{
					_context.AddToken(CreateToken(_context, Worksheet));
				}
				if (token.TokenType == TokenType.OpeningParenthesis && (_context.LastToken.TokenType == TokenType.ExcelAddress || _context.LastToken.TokenType == TokenType.NameValue))
				{
					_context.LastToken.TokenType = TokenType.Function;
				}
			}
			if (token.Value == "-" && TokenIsNegator(_context))
			{
				_context.AddToken(new Token("-", TokenType.Negator));
				return;
			}
			_context.AddToken(token);
			_context.NewToken();
		}
		else
		{
			_context.AppendToCurrentToken(c);
		}
	}

	private bool CharIsTokenSeparator(char c, out Token token)
	{
		bool flag = _tokenProvider.Tokens.ContainsKey(c.ToString());
		token = (flag ? (token = _tokenProvider.Tokens[c.ToString()]) : null);
		return flag;
	}

	private static bool TokenIsNegator(TokenizerContext context)
	{
		return TokenIsNegator(context.LastToken);
	}

	private static bool TokenIsNegator(Token t)
	{
		if (t != null && t.TokenType != 0 && t.TokenType != TokenType.OpeningParenthesis && t.TokenType != TokenType.Comma && t.TokenType != TokenType.SemiColon)
		{
			return t.TokenType == TokenType.OpeningEnumerable;
		}
		return true;
	}

	private Token CreateToken(TokenizerContext context, string worksheet)
	{
		if (context.CurrentToken == "-" && context.LastToken == null && context.LastToken.TokenType == TokenType.Operator)
		{
			return new Token("-", TokenType.Negator);
		}
		return _tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
	}

	void ITokenIndexProvider.MoveIndexPointerForward()
	{
		_tokenIndex++;
	}
}
