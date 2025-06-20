namespace OfficeOpenXml.FormulaParsing;

public class ExcelCalculationOption
{
	public bool AllowCirculareReferences { get; set; }

	public ExcelCalculationOption()
	{
		AllowCirculareReferences = false;
	}
}
