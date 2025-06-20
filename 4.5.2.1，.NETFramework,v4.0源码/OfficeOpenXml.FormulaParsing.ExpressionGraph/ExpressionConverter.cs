using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExpressionConverter : IExpressionConverter
{
	private static IExpressionConverter _instance;

	public static IExpressionConverter Instance
	{
		get
		{
			if (_instance == null)
			{
				_instance = new ExpressionConverter();
			}
			return _instance;
		}
	}

	public StringExpression ToStringExpression(Expression expression)
	{
		return new StringExpression(expression.Compile().Result.ToString())
		{
			Operator = expression.Operator
		};
	}

	public Expression FromCompileResult(CompileResult compileResult)
	{
		switch (compileResult.DataType)
		{
		case DataType.Integer:
			if (!(compileResult.Result is string))
			{
				return new IntegerExpression(Convert.ToDouble(compileResult.Result));
			}
			return new IntegerExpression(compileResult.Result.ToString());
		case DataType.String:
			return new StringExpression(compileResult.Result.ToString());
		case DataType.Decimal:
			if (!(compileResult.Result is string))
			{
				return new DecimalExpression((double)compileResult.Result);
			}
			return new DecimalExpression(compileResult.Result.ToString());
		case DataType.Boolean:
			if (!(compileResult.Result is string))
			{
				return new BooleanExpression((bool)compileResult.Result);
			}
			return new BooleanExpression(compileResult.Result.ToString());
		case DataType.ExcelError:
			if (!(compileResult.Result is string))
			{
				return new ExcelErrorExpression((ExcelErrorValue)compileResult.Result);
			}
			return new ExcelErrorExpression(compileResult.Result.ToString(), ExcelErrorValue.Parse(compileResult.Result.ToString()));
		case DataType.Empty:
			return new IntegerExpression(0.0);
		default:
			return null;
		}
	}
}
