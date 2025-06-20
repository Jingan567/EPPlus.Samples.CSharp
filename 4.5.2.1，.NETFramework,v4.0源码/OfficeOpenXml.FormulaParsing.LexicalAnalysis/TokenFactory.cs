using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class TokenFactory : ITokenFactory
{
	private readonly ITokenSeparatorProvider _tokenSeparatorProvider;

	private readonly IFunctionNameProvider _functionNameProvider;

	private readonly INameValueProvider _nameValueProvider;

	private bool _r1c1;

	public TokenFactory(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1 = false)
		: this(new TokenSeparatorProvider(), nameValueProvider, functionRepository, r1c1)
	{
	}

	public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider, INameValueProvider nameValueProvider, IFunctionNameProvider functionNameProvider, bool r1c1)
	{
		_tokenSeparatorProvider = tokenSeparatorProvider;
		_functionNameProvider = functionNameProvider;
		_nameValueProvider = nameValueProvider;
		_r1c1 = r1c1;
	}

	public Token Create(IEnumerable<Token> tokens, string token)
	{
		return Create(tokens, token, null);
	}

	public Token Create(IEnumerable<Token> tokens, string token, string worksheet)
	{
		Token value = null;
		if (_tokenSeparatorProvider.Tokens.TryGetValue(token, out value))
		{
			return value;
		}
		IList<Token> list = (IList<Token>)tokens;
		if (token.StartsWith("!") && list[list.Count - 1].TokenType == TokenType.String)
		{
			string text = "";
			int num = list.Count - 2;
			if (num > 0)
			{
				if (list[num].TokenType == TokenType.StringContent)
				{
					text = "'" + list[num].Value.Replace("'", "''") + "'";
					list.RemoveAt(list.Count - 1);
					list.RemoveAt(list.Count - 1);
					list.RemoveAt(list.Count - 1);
					return new Token(text + token, TokenType.ExcelAddress);
				}
				throw new ArgumentException($"Invalid formula token sequence near {token}");
			}
			throw new ArgumentException($"Invalid formula token sequence near {token}");
		}
		if (tokens.Any() && tokens.Last().TokenType == TokenType.String)
		{
			return new Token(token, TokenType.StringContent);
		}
		if (!string.IsNullOrEmpty(token))
		{
			token = token.Trim();
		}
		if (Regex.IsMatch(token, "^[0-9]+\\.[0-9]+$"))
		{
			return new Token(token, TokenType.Decimal);
		}
		if (Regex.IsMatch(token, "^[0-9]+$"))
		{
			return new Token(token, TokenType.Integer);
		}
		if (Regex.IsMatch(token, "^(true|false)$", RegexOptions.IgnoreCase))
		{
			return new Token(token, TokenType.Boolean);
		}
		if (token.ToUpper(CultureInfo.InvariantCulture).Contains("#REF!"))
		{
			return new Token(token, TokenType.InvalidReference);
		}
		if (token.ToUpper(CultureInfo.InvariantCulture) == "#NUM!")
		{
			return new Token(token, TokenType.NumericError);
		}
		if (token.ToUpper(CultureInfo.InvariantCulture) == "#VALUE!")
		{
			return new Token(token, TokenType.ValueDataTypeError);
		}
		if (token.ToUpper(CultureInfo.InvariantCulture) == "#NULL!")
		{
			return new Token(token, TokenType.Null);
		}
		if (_nameValueProvider != null && _nameValueProvider.IsNamedValue(token, worksheet))
		{
			return new Token(token, TokenType.NameValue);
		}
		if (_functionNameProvider.IsFunctionName(token))
		{
			return new Token(token, TokenType.Function);
		}
		if (list.Count > 0 && list[list.Count - 1].TokenType == TokenType.OpeningEnumerable)
		{
			return new Token(token, TokenType.Enumerable);
		}
		return ExcelAddressBase.IsValid(token, _r1c1) switch
		{
			ExcelAddressBase.AddressType.InternalAddress => new Token(token.ToUpper(CultureInfo.InvariantCulture), TokenType.ExcelAddress), 
			ExcelAddressBase.AddressType.R1C1 => new Token(token.ToUpper(CultureInfo.InvariantCulture), TokenType.ExcelAddressR1C1), 
			_ => new Token(token, TokenType.Unrecognized), 
		};
	}

	public Token Create(string token, TokenType explicitTokenType)
	{
		return new Token(token, explicitTokenType);
	}
}
