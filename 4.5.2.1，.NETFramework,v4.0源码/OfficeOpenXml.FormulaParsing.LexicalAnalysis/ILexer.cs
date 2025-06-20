using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public interface ILexer
{
	IEnumerable<Token> Tokenize(string input);

	IEnumerable<Token> Tokenize(string input, string worksheet);
}
