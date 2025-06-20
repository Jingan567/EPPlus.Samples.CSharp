using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public interface ITokenFactory
{
	Token Create(IEnumerable<Token> tokens, string token);

	Token Create(IEnumerable<Token> tokens, string token, string worksheet);

	Token Create(string token, TokenType explicitTokenType);
}
