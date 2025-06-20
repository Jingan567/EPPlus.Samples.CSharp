namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class FormulaDependencyFactory
{
	public virtual FormulaDependency Create(ParsingScope scope)
	{
		return new FormulaDependency(scope);
	}
}
