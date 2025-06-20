namespace OfficeOpenXml.Compatibility;

public class CompatibilitySettings
{
	private ExcelPackage excelPackage;

	public bool IsWorksheets1Based
	{
		get
		{
			return excelPackage._worksheetAdd == 1;
		}
		set
		{
			excelPackage._worksheetAdd = (value ? 1 : 0);
			if (excelPackage._workbook != null && excelPackage._workbook._worksheets != null)
			{
				excelPackage.Workbook.Worksheets.ReindexWorksheetDictionary();
			}
		}
	}

	internal CompatibilitySettings(ExcelPackage excelPackage)
	{
		this.excelPackage = excelPackage;
	}
}
