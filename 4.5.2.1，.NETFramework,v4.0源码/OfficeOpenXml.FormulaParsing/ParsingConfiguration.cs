using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing;

public class ParsingConfiguration
{
	public virtual ILexer Lexer { get; private set; }

	public IFormulaParserLogger Logger { get; private set; }

	public IExpressionGraphBuilder GraphBuilder { get; private set; }

	public IExpressionCompiler ExpressionCompiler { get; private set; }

	public FunctionRepository FunctionRepository { get; private set; }

	private ParsingConfiguration()
	{
		FunctionRepository = FunctionRepository.Create();
	}

	internal static ParsingConfiguration Create()
	{
		return new ParsingConfiguration();
	}

	public ParsingConfiguration SetLexer(ILexer lexer)
	{
		Lexer = lexer;
		return this;
	}

	public ParsingConfiguration SetGraphBuilder(IExpressionGraphBuilder graphBuilder)
	{
		GraphBuilder = graphBuilder;
		return this;
	}

	public ParsingConfiguration SetExpresionCompiler(IExpressionCompiler expressionCompiler)
	{
		ExpressionCompiler = expressionCompiler;
		return this;
	}

	public ParsingConfiguration AttachLogger(IFormulaParserLogger logger)
	{
		Require.That(logger).Named("logger").IsNotNull();
		Logger = logger;
		return this;
	}

	public ParsingConfiguration DetachLogger()
	{
		Logger = null;
		return this;
	}
}
