using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing;

public class FormulaParser : IDisposable
{
	private readonly ParsingContext _parsingContext;

	private readonly ExcelDataProvider _excelDataProvider;

	private ILexer _lexer;

	private IExpressionGraphBuilder _graphBuilder;

	private IExpressionCompiler _compiler;

	public ILexer Lexer => _lexer;

	public IEnumerable<string> FunctionNames => _parsingContext.Configuration.FunctionRepository.FunctionNames;

	public IFormulaParserLogger Logger => _parsingContext.Configuration.Logger;

	public FormulaParser(ExcelDataProvider excelDataProvider)
		: this(excelDataProvider, ParsingContext.Create())
	{
	}

	public FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
	{
		FormulaParser formulaParser = this;
		parsingContext.Parser = this;
		parsingContext.ExcelDataProvider = excelDataProvider;
		parsingContext.NameValueProvider = new EpplusNameValueProvider(excelDataProvider);
		parsingContext.RangeAddressFactory = new RangeAddressFactory(excelDataProvider);
		_parsingContext = parsingContext;
		_excelDataProvider = excelDataProvider;
		Configure(delegate(ParsingConfiguration configuration)
		{
			configuration.SetLexer(new Lexer(formulaParser._parsingContext.Configuration.FunctionRepository, formulaParser._parsingContext.NameValueProvider)).SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, formulaParser._parsingContext)).SetExpresionCompiler(new ExpressionCompiler())
				.FunctionRepository.LoadModule(new BuiltInFunctions());
		});
	}

	public void Configure(Action<ParsingConfiguration> configMethod)
	{
		configMethod(_parsingContext.Configuration);
		_lexer = _parsingContext.Configuration.Lexer ?? _lexer;
		_graphBuilder = _parsingContext.Configuration.GraphBuilder ?? _graphBuilder;
		_compiler = _parsingContext.Configuration.ExpressionCompiler ?? _compiler;
	}

	internal virtual object Parse(string formula, RangeAddress rangeAddress)
	{
		using (_parsingContext.Scopes.NewScope(rangeAddress))
		{
			IEnumerable<Token> tokens = _lexer.Tokenize(formula);
			OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph expressionGraph = _graphBuilder.Build(tokens);
			if (expressionGraph.Expressions.Count() == 0)
			{
				return null;
			}
			return _compiler.Compile(expressionGraph.Expressions).Result;
		}
	}

	internal virtual object Parse(IEnumerable<Token> tokens, string worksheet, string address)
	{
		RangeAddress address2 = _parsingContext.RangeAddressFactory.Create(address);
		using (_parsingContext.Scopes.NewScope(address2))
		{
			OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph expressionGraph = _graphBuilder.Build(tokens);
			if (expressionGraph.Expressions.Count() == 0)
			{
				return null;
			}
			return _compiler.Compile(expressionGraph.Expressions).Result;
		}
	}

	internal virtual object ParseCell(IEnumerable<Token> tokens, string worksheet, int row, int column)
	{
		RangeAddress address = _parsingContext.RangeAddressFactory.Create(worksheet, column, row);
		using (_parsingContext.Scopes.NewScope(address))
		{
			OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph expressionGraph = _graphBuilder.Build(tokens);
			if (expressionGraph.Expressions.Count() == 0)
			{
				return 0.0;
			}
			try
			{
				CompileResult compileResult = _compiler.Compile(expressionGraph.Expressions);
				if (!(compileResult.Result is ExcelDataProvider.IRangeInfo rangeInfo))
				{
					return compileResult.Result ?? ((object)0.0);
				}
				if (rangeInfo.IsEmpty)
				{
					return 0.0;
				}
				if (!rangeInfo.IsMulti)
				{
					return rangeInfo.First().Value ?? ((object)0.0);
				}
				if (string.IsNullOrEmpty(worksheet))
				{
					return rangeInfo;
				}
				if (_parsingContext.Debug)
				{
					string message = $"A range with multiple cell was returned at row {row}, column {column}";
					_parsingContext.Configuration.Logger.Log(_parsingContext, message);
				}
				return ExcelErrorValue.Create(eErrorType.Value);
			}
			catch (ExcelErrorValueException ex)
			{
				if (_parsingContext.Debug)
				{
					_parsingContext.Configuration.Logger.Log(_parsingContext, ex);
				}
				return ex.ErrorValue;
			}
		}
	}

	public virtual object Parse(string formula, string address)
	{
		return Parse(formula, _parsingContext.RangeAddressFactory.Create(address));
	}

	public virtual object Parse(string formula)
	{
		return Parse(formula, RangeAddress.Empty);
	}

	public virtual object ParseAt(string address)
	{
		Require.That(address).Named("address").IsNotNullOrEmpty();
		RangeAddress rangeAddress = _parsingContext.RangeAddressFactory.Create(address);
		return ParseAt(rangeAddress.Worksheet, rangeAddress.FromRow, rangeAddress.FromCol);
	}

	public virtual object ParseAt(string worksheetName, int row, int col)
	{
		string rangeFormula = _excelDataProvider.GetRangeFormula(worksheetName, row, col);
		if (string.IsNullOrEmpty(rangeFormula))
		{
			return _excelDataProvider.GetRangeValue(worksheetName, row, col);
		}
		return Parse(rangeFormula, _parsingContext.RangeAddressFactory.Create(worksheetName, col, row));
	}

	internal void InitNewCalc()
	{
		if (_excelDataProvider != null)
		{
			_excelDataProvider.Reset();
		}
	}

	public void Dispose()
	{
		if (_parsingContext.Debug)
		{
			_parsingContext.Configuration.Logger.Dispose();
		}
	}
}
