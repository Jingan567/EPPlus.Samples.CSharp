using System;

namespace OfficeOpenXml.Style;

internal interface IExcelCell
{
	object Value { get; set; }

	string StyleName { get; }

	int StyleID { get; set; }

	ExcelStyle Style { get; }

	Uri Hyperlink { get; set; }

	string Formula { get; set; }

	string FormulaR1C1 { get; set; }
}
