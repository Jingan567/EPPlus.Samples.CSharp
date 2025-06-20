using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public interface ISourceCodeTokenizer
{
	IEnumerable<Token> Tokenize(string input, string worksheet);
}
