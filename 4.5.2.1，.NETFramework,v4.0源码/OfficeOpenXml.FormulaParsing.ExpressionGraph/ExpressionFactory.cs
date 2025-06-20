using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExpressionFactory : IExpressionFactory
{
	private readonly ExcelDataProvider _excelDataProvider;

	private readonly ParsingContext _parsingContext;

	public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
	{
		_excelDataProvider = excelDataProvider;
		_parsingContext = context;
	}

	public Expression Create(Token token)
	{
		return token.TokenType switch
		{
			TokenType.Integer => new IntegerExpression(token.Value, token.IsNegated), 
			TokenType.String => new StringExpression(token.Value), 
			TokenType.Decimal => new DecimalExpression(token.Value, token.IsNegated), 
			TokenType.Boolean => new BooleanExpression(token.Value), 
			TokenType.ExcelAddress => new ExcelAddressExpression(token.Value, _excelDataProvider, _parsingContext, token.IsNegated), 
			TokenType.InvalidReference => new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Ref)), 
			TokenType.NumericError => new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Num)), 
			TokenType.ValueDataTypeError => new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Value)), 
			TokenType.Null => new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Null)), 
			TokenType.NameValue => new NamedValueExpression(token.Value, _parsingContext), 
			_ => new StringExpression(token.Value), 
		};
	}
}
