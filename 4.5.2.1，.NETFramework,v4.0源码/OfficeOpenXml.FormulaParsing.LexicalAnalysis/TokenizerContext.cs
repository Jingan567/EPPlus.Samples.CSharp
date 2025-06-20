using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class TokenizerContext
{
	private char[] _chars;

	private List<Token> _result;

	private StringBuilder _currentToken;

	public char[] FormulaChars => _chars;

	public IList<Token> Result => _result;

	public bool IsInString { get; private set; }

	public bool IsInSheetName { get; private set; }

	internal int BracketCount { get; set; }

	public string CurrentToken => _currentToken.ToString();

	public bool CurrentTokenHasValue => !string.IsNullOrEmpty(IsInString ? CurrentToken : CurrentToken.Trim());

	public Token LastToken
	{
		get
		{
			if (_result.Count <= 0)
			{
				return null;
			}
			return _result.Last();
		}
	}

	public TokenizerContext(string formula)
	{
		if (!string.IsNullOrEmpty(formula))
		{
			_chars = formula.ToArray();
		}
		_result = new List<Token>();
		_currentToken = new StringBuilder();
	}

	public void ToggleIsInString()
	{
		IsInString = !IsInString;
	}

	public void ToggleIsInSheetName()
	{
		IsInSheetName = !IsInSheetName;
	}

	public void NewToken()
	{
		_currentToken = new StringBuilder();
	}

	public void AddToken(Token token)
	{
		_result.Add(token);
	}

	public void AppendToCurrentToken(char c)
	{
		_currentToken.Append(c.ToString());
	}

	public void AppendToLastToken(string stringToAppend)
	{
		_result.Last().Append(stringToAppend);
	}

	public void SetLastTokenType(TokenType type)
	{
		_result.Last().TokenType = type;
	}

	public void ReplaceLastToken(Token newToken)
	{
		int count = _result.Count;
		if (count > 0)
		{
			_result.RemoveAt(count - 1);
		}
		_result.Add(newToken);
	}
}
