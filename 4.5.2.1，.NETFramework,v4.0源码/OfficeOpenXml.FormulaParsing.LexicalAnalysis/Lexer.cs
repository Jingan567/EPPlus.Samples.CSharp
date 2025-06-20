using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class Lexer : ILexer
{
	private readonly ISourceCodeTokenizer _tokenizer;

	private readonly ISyntacticAnalyzer _analyzer;

	public Lexer(FunctionRepository functionRepository, INameValueProvider nameValueProvider)
		: this(new SourceCodeTokenizer(functionRepository, nameValueProvider), new SyntacticAnalyzer())
	{
	}

	public Lexer(ISourceCodeTokenizer tokenizer, ISyntacticAnalyzer analyzer)
	{
		_tokenizer = tokenizer;
		_analyzer = analyzer;
	}

	public IEnumerable<Token> Tokenize(string input)
	{
		return Tokenize(input, null);
	}

	public IEnumerable<Token> Tokenize(string input, string worksheet)
	{
		IEnumerable<Token> enumerable = _tokenizer.Tokenize(input, worksheet);
		_analyzer.Analyze(enumerable);
		return enumerable;
	}
}
