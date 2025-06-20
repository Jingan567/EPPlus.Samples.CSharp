namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class FunctionNameProvider : IFunctionNameProvider
{
	public static FunctionNameProvider Empty => new FunctionNameProvider();

	private FunctionNameProvider()
	{
	}

	public virtual bool IsFunctionName(string name)
	{
		return false;
	}
}
