namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class Token
{
	public string Value { get; internal set; }

	public TokenType TokenType { get; internal set; }

	public bool IsNegated { get; private set; }

	public Token(string token, TokenType tokenType)
	{
		Value = token;
		TokenType = tokenType;
	}

	public void Append(string stringToAppend)
	{
		Value += stringToAppend;
	}

	public void Negate()
	{
		if (TokenType == TokenType.Decimal || TokenType == TokenType.Integer || TokenType == TokenType.ExcelAddress)
		{
			IsNegated = true;
		}
	}

	public override string ToString()
	{
		return TokenType.ToString() + ", " + Value;
	}
}
