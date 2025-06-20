using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public interface ISyntacticAnalyzer
{
	void Analyze(IEnumerable<Token> tokens);
}
