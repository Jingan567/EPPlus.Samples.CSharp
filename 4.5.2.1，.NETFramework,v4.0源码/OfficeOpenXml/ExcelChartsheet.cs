using System;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;

namespace OfficeOpenXml;

public class ExcelChartsheet : ExcelWorksheet
{
	public ExcelChart Chart => (ExcelChart)base.Drawings[0];

	public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden, eChartType chartType)
		: base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
	{
		base.Drawings.AddChart("Chart 1", chartType);
	}

	public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden)
		: base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
	{
	}
}
