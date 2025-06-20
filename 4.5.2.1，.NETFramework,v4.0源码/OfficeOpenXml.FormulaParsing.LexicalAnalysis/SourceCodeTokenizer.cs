using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class SourceCodeTokenizer : ISourceCodeTokenizer
{
	private readonly ITokenSeparatorProvider _separatorProvider;

	private readonly ITokenFactory _tokenFactory;

	public static ISourceCodeTokenizer Default => new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);

	public static ISourceCodeTokenizer R1C1 => new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, r1c1: true);

	public SourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1 = false)
		: this(new TokenFactory(functionRepository, nameValueProvider, r1c1), new TokenSeparatorProvider())
	{
	}

	public SourceCodeTokenizer(ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
	{
		_tokenFactory = tokenFactory;
		_separatorProvider = tokenProvider;
	}

	public IEnumerable<Token> Tokenize(string input)
	{
		return Tokenize(input, null);
	}

	public IEnumerable<Token> Tokenize(string input, string worksheet)
	{
		if (string.IsNullOrEmpty(input))
		{
			return Enumerable.Empty<Token>();
		}
		input = input.TrimStart('+');
		TokenizerContext tokenizerContext = new TokenizerContext(input);
		TokenHandler tokenHandler = new TokenHandler(tokenizerContext, _tokenFactory, _separatorProvider);
		tokenHandler.Worksheet = worksheet;
		while (tokenHandler.HasMore())
		{
			tokenHandler.Next();
		}
		if (tokenizerContext.CurrentTokenHasValue)
		{
			tokenizerContext.AddToken(CreateToken(tokenizerContext, worksheet));
		}
		CleanupTokens(tokenizerContext, _separatorProvider.Tokens);
		return tokenizerContext.Result;
	}

	private static void CleanupTokens(TokenizerContext context, IDictionary<string, Token> tokens)
	{
		for (int i = 0; i < context.Result.Count; i++)
		{
			Token token = context.Result[i];
			if (token.TokenType == TokenType.Unrecognized)
			{
				if (i < context.Result.Count - 1)
				{
					if (context.Result[i + 1].TokenType == TokenType.OpeningParenthesis)
					{
						token.TokenType = TokenType.Function;
					}
					else
					{
						token.TokenType = TokenType.NameValue;
					}
				}
				else
				{
					token.TokenType = TokenType.NameValue;
				}
			}
			else if (token.TokenType == TokenType.WorksheetName)
			{
				token.TokenType = context.Result[i + 3].TokenType;
				StringBuilder stringBuilder = new StringBuilder();
				int num = 3;
				if (context.Result.Count < i + num)
				{
					token.TokenType = TokenType.InvalidReference;
					num = context.Result.Count - i - 1;
				}
				else if (context.Result[i + 3].TokenType != TokenType.ExcelAddress && context.Result[i + 3].TokenType != TokenType.ExcelAddressR1C1)
				{
					token.TokenType = TokenType.InvalidReference;
					num--;
				}
				else
				{
					for (int j = 0; j < 4; j++)
					{
						stringBuilder.Append(context.Result[i + j].Value);
					}
				}
				token.Value = stringBuilder.ToString();
				for (int k = 0; k < num; k++)
				{
					context.Result.RemoveAt(i + 1);
				}
			}
			else
			{
				if ((token.TokenType != 0 && token.TokenType != TokenType.Negator) || i >= context.Result.Count - 1 || (!(token.Value == "+") && !(token.Value == "-")))
				{
					continue;
				}
				if (token.Value == "+" && (i == 0 || context.Result[i - 1].TokenType == TokenType.OpeningParenthesis || context.Result[i - 1].TokenType == TokenType.Comma))
				{
					context.Result.RemoveAt(i);
					SetNegatorOperator(context, i, tokens);
					i--;
					continue;
				}
				Token token2 = context.Result[i + 1];
				if (token2.TokenType == TokenType.Operator || token2.TokenType == TokenType.Negator)
				{
					if (token.Value == "+" && (token2.Value == "+" || token2.Value == "-"))
					{
						context.Result.RemoveAt(i);
						SetNegatorOperator(context, i, tokens);
						i--;
					}
					else if (token.Value == "-" && token2.Value == "+")
					{
						context.Result.RemoveAt(i + 1);
						SetNegatorOperator(context, i, tokens);
						i--;
					}
					else if (token.Value == "-" && token2.Value == "-")
					{
						context.Result.RemoveAt(i);
						context.Result[i] = tokens["+"];
						i--;
					}
				}
			}
		}
	}

	private static void SetNegatorOperator(TokenizerContext context, int i, IDictionary<string, Token> tokens)
	{
		if (context.Result[i].Value == "-" && i > 0 && (context.Result[i].TokenType == TokenType.Operator || context.Result[i].TokenType == TokenType.Negator))
		{
			if (TokenIsNegator(context.Result[i - 1]))
			{
				context.Result[i] = new Token("-", TokenType.Negator);
			}
			else
			{
				context.Result[i] = tokens["-"];
			}
		}
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
}
