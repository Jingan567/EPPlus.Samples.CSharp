using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class FunctionExpression : AtomicExpression
{
	private readonly ParsingContext _parsingContext;

	private readonly FunctionCompilerFactory _functionCompilerFactory;

	private readonly bool _isNegated;

	public override bool HasChildren
	{
		get
		{
			if (base.Children.Any())
			{
				return base.Children.First().Children.Any();
			}
			return false;
		}
	}

	public FunctionExpression(string expression, ParsingContext parsingContext, bool isNegated)
		: base(expression)
	{
		_parsingContext = parsingContext;
		_functionCompilerFactory = new FunctionCompilerFactory(parsingContext.Configuration.FunctionRepository);
		_isNegated = isNegated;
		base.AddChild(new FunctionArgumentExpression(this));
	}

	public override CompileResult Compile()
	{
		try
		{
			ExcelFunction function = _parsingContext.Configuration.FunctionRepository.GetFunction(base.ExpressionString);
			if (function == null)
			{
				if (_parsingContext.Debug)
				{
					_parsingContext.Configuration.Logger.Log(_parsingContext, $"'{base.ExpressionString}' is not a supported function");
				}
				return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);
			}
			if (_parsingContext.Debug)
			{
				_parsingContext.Configuration.Logger.LogFunction(base.ExpressionString);
			}
			CompileResult compileResult = _functionCompilerFactory.Create(function).Compile(HasChildren ? base.Children : Enumerable.Empty<Expression>(), _parsingContext);
			if (_isNegated)
			{
				if (!compileResult.IsNumeric)
				{
					if (_parsingContext.Debug)
					{
						string message = $"Trying to negate a non-numeric value ({compileResult.Result}) in function '{base.ExpressionString}'";
						_parsingContext.Configuration.Logger.Log(_parsingContext, message);
					}
					return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
				}
				return new CompileResult(compileResult.ResultNumeric * -1.0, compileResult.DataType);
			}
			return compileResult;
		}
		catch (ExcelErrorValueException ex)
		{
			if (_parsingContext.Debug)
			{
				_parsingContext.Configuration.Logger.Log(_parsingContext, ex);
			}
			return new CompileResult(ex.ErrorValue, DataType.ExcelError);
		}
	}

	public override Expression PrepareForNextChild()
	{
		return base.AddChild(new FunctionArgumentExpression(this));
	}

	public override Expression AddChild(Expression child)
	{
		base.Children.Last().AddChild(child);
		return child;
	}
}
