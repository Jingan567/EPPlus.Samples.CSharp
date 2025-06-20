namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public abstract class CompileResultValidator
{
	private static CompileResultValidator _empty;

	public static CompileResultValidator Empty => _empty ?? (_empty = new EmptyCompileResultValidator());

	public abstract void Validate(object obj);
}
