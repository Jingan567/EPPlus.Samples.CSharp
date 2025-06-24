namespace OfficeOpenXml.Compatibility;

public class CompatibilitySettings
{
	private ExcelPackage excelPackage;

    /// <summary>
    /// 默认情况下，工作表索引从 0 开始。
	/// 在 .NET Framework 中，您可以通过设置 Compatibility.IsWorksheets1Based = true 来更改此行为。
    /// </summary>
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
