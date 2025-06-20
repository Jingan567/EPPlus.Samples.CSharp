using System;
using System.Collections.Generic;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public static class CalculationExtension
{
	public static void Calculate(this ExcelWorkbook workbook)
	{
		workbook.Calculate(new ExcelCalculationOption
		{
			AllowCirculareReferences = false
		});
	}

	public static void Calculate(this ExcelWorkbook workbook, ExcelCalculationOption options)
	{
		Init(workbook);
		DependencyChain dependencyChain = DependencyChainFactory.Create(workbook, options);
		workbook.FormulaParser.InitNewCalc();
		if (workbook.FormulaParser.Logger != null)
		{
			string message = $"Starting... number of cells to parse: {dependencyChain.list.Count}";
			workbook.FormulaParser.Logger.Log(message);
		}
		CalcChain(workbook, workbook.FormulaParser, dependencyChain);
	}

	public static void Calculate(this ExcelWorksheet worksheet)
	{
		worksheet.Calculate(new ExcelCalculationOption());
	}

	public static void Calculate(this ExcelWorksheet worksheet, ExcelCalculationOption options)
	{
		Init(worksheet.Workbook);
		DependencyChain dependencyChain = DependencyChainFactory.Create(worksheet, options);
		FormulaParser formulaParser = worksheet.Workbook.FormulaParser;
		formulaParser.InitNewCalc();
		if (formulaParser.Logger != null)
		{
			string message = $"Starting... number of cells to parse: {dependencyChain.list.Count}";
			formulaParser.Logger.Log(message);
		}
		CalcChain(worksheet.Workbook, formulaParser, dependencyChain);
	}

	public static void Calculate(this ExcelRangeBase range)
	{
		range.Calculate(new ExcelCalculationOption());
	}

	public static void Calculate(this ExcelRangeBase range, ExcelCalculationOption options)
	{
		Init(range._workbook);
		FormulaParser formulaParser = range._workbook.FormulaParser;
		formulaParser.InitNewCalc();
		DependencyChain dc = DependencyChainFactory.Create(range, options);
		CalcChain(range._workbook, formulaParser, dc);
	}

	public static object Calculate(this ExcelWorksheet worksheet, string Formula)
	{
		return worksheet.Calculate(Formula, new ExcelCalculationOption());
	}

	public static object Calculate(this ExcelWorksheet worksheet, string Formula, ExcelCalculationOption options)
	{
		try
		{
			worksheet.CheckSheetType();
			if (string.IsNullOrEmpty(Formula.Trim()))
			{
				return null;
			}
			Init(worksheet.Workbook);
			FormulaParser formulaParser = worksheet.Workbook.FormulaParser;
			formulaParser.InitNewCalc();
			if (Formula[0] == '=')
			{
				Formula = Formula.Substring(1);
			}
			DependencyChain dependencyChain = DependencyChainFactory.Create(worksheet, Formula, options);
			FormulaCell formulaCell = dependencyChain.list[0];
			dependencyChain.CalcOrder.RemoveAt(dependencyChain.CalcOrder.Count - 1);
			CalcChain(worksheet.Workbook, formulaParser, dependencyChain);
			return formulaParser.ParseCell(formulaCell.Tokens, worksheet.Name, -1, -1);
		}
		catch (Exception ex)
		{
			return new ExcelErrorValueException(ex.Message, ExcelErrorValue.Create(eErrorType.Value));
		}
	}

	private static void CalcChain(ExcelWorkbook wb, FormulaParser parser, DependencyChain dc)
	{
		bool flag = parser.Logger != null;
		foreach (int item in dc.CalcOrder)
		{
			FormulaCell formulaCell = dc.list[item];
			try
			{
				ExcelWorksheet bySheetID = wb.Worksheets.GetBySheetID(formulaCell.SheetID);
				object v = parser.ParseCell(formulaCell.Tokens, (bySheetID == null) ? "" : bySheetID.Name, formulaCell.Row, formulaCell.Column);
				SetValue(wb, formulaCell, v);
				if (flag)
				{
					parser.Logger.LogCellCounted();
				}
				Thread.Sleep(0);
			}
			catch (FormatException ex)
			{
				throw ex;
			}
			catch
			{
				ExcelErrorValue v2 = ExcelErrorValue.Parse("#VALUE!");
				SetValue(wb, formulaCell, v2);
			}
		}
	}

	private static void Init(ExcelWorkbook workbook)
	{
		workbook._formulaTokens = new CellStore<List<Token>>();
		foreach (ExcelWorksheet worksheet in workbook.Worksheets)
		{
			if (!(worksheet is ExcelChartsheet))
			{
				if (worksheet._formulaTokens != null)
				{
					worksheet._formulaTokens.Dispose();
				}
				worksheet._formulaTokens = new CellStore<List<Token>>();
			}
		}
	}

	private static void SetValue(ExcelWorkbook workbook, FormulaCell item, object v)
	{
		if (item.Column == 0)
		{
			if (item.SheetID <= 0)
			{
				workbook.Names[item.Row].NameValue = v;
			}
			else
			{
				workbook.Worksheets.GetBySheetID(item.SheetID).Names[item.Row].NameValue = v;
			}
		}
		else
		{
			workbook.Worksheets.GetBySheetID(item.SheetID).SetValueInner(item.Row, item.Column, v);
		}
	}
}
