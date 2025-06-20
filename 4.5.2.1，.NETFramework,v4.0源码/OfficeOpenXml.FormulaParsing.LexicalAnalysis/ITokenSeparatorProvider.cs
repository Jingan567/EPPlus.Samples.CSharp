using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public interface ITokenSeparatorProvider
{
	IDictionary<string, Token> Tokens { get; }

	bool IsOperator(string item);

	bool IsPossibleLastPartOfMultipleCharOperator(string part);
}
