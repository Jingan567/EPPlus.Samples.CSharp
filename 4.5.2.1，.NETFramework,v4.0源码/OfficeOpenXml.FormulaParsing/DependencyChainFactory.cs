using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing;

internal static class DependencyChainFactory
{
	internal static DependencyChain Create(ExcelWorkbook wb, ExcelCalculationOption options)
	{
		DependencyChain dependencyChain = new DependencyChain();
		foreach (ExcelWorksheet worksheet in wb.Worksheets)
		{
			if (!(worksheet is ExcelChartsheet))
			{
				GetChain(dependencyChain, wb.FormulaParser.Lexer, worksheet.Cells, options);
				GetWorksheetNames(worksheet, dependencyChain, options);
			}
		}
		foreach (ExcelNamedRange name in wb.Names)
		{
			if (name.NameValue == null)
			{
				GetChain(dependencyChain, wb.FormulaParser.Lexer, name, options);
			}
		}
		return dependencyChain;
	}

	internal static DependencyChain Create(ExcelWorksheet ws, ExcelCalculationOption options)
	{
		ws.CheckSheetType();
		DependencyChain dependencyChain = new DependencyChain();
		GetChain(dependencyChain, ws.Workbook.FormulaParser.Lexer, ws.Cells, options);
		GetWorksheetNames(ws, dependencyChain, options);
		return dependencyChain;
	}

	internal static DependencyChain Create(ExcelWorksheet ws, string Formula, ExcelCalculationOption options)
	{
		ws.CheckSheetType();
		DependencyChain dependencyChain = new DependencyChain();
		GetChain(dependencyChain, ws.Workbook.FormulaParser.Lexer, ws, Formula, options);
		return dependencyChain;
	}

	private static void GetWorksheetNames(ExcelWorksheet ws, DependencyChain depChain, ExcelCalculationOption options)
	{
		foreach (ExcelNamedRange name in ws.Names)
		{
			if (!string.IsNullOrEmpty(name.NameFormula))
			{
				GetChain(depChain, ws.Workbook.FormulaParser.Lexer, name, options);
			}
		}
	}

	internal static DependencyChain Create(ExcelRangeBase range, ExcelCalculationOption options)
	{
		DependencyChain dependencyChain = new DependencyChain();
		GetChain(dependencyChain, range.Worksheet.Workbook.FormulaParser.Lexer, range, options);
		return dependencyChain;
	}

	private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelNamedRange name, ExcelCalculationOption options)
	{
		ExcelWorksheet worksheet = name.Worksheet;
		ulong cellID = ExcelCellBase.GetCellID(worksheet?.SheetID ?? 0, name.Index, 0);
		if (depChain.index.ContainsKey(cellID))
		{
			return;
		}
		FormulaCell formulaCell = new FormulaCell
		{
			SheetID = (worksheet?.SheetID ?? 0),
			Row = name.Index,
			Column = 0,
			Formula = name.NameFormula
		};
		if (!string.IsNullOrEmpty(formulaCell.Formula))
		{
			formulaCell.Tokens = lexer.Tokenize(formulaCell.Formula, worksheet?.Name).ToList();
			if (worksheet == null)
			{
				name._workbook._formulaTokens.SetValue(name.Index, 0, formulaCell.Tokens);
			}
			else
			{
				worksheet._formulaTokens.SetValue(name.Index, 0, formulaCell.Tokens);
			}
			depChain.Add(formulaCell);
			FollowChain(depChain, lexer, name._workbook, worksheet, formulaCell, options);
		}
	}

	private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelWorksheet ws, string formula, ExcelCalculationOption options)
	{
		FormulaCell formulaCell = new FormulaCell
		{
			SheetID = ws.SheetID,
			Row = -1,
			Column = -1
		};
		formulaCell.Formula = formula;
		if (!string.IsNullOrEmpty(formulaCell.Formula))
		{
			formulaCell.Tokens = lexer.Tokenize(formulaCell.Formula, ws.Name).ToList();
			depChain.Add(formulaCell);
			FollowChain(depChain, lexer, ws.Workbook, ws, formulaCell, options);
		}
	}

	private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelRangeBase Range, ExcelCalculationOption options)
	{
		ExcelWorksheet worksheet = Range.Worksheet;
		CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(worksheet._formulas, Range.Start.Row, Range.Start.Column, Range.End.Row, Range.End.Column);
		while (cellsStoreEnumerator.Next())
		{
			if (cellsStoreEnumerator.Value == null || cellsStoreEnumerator.Value.ToString().Trim() == "")
			{
				continue;
			}
			ulong cellID = ExcelCellBase.GetCellID(worksheet.SheetID, cellsStoreEnumerator.Row, cellsStoreEnumerator.Column);
			if (!depChain.index.ContainsKey(cellID))
			{
				FormulaCell formulaCell = new FormulaCell
				{
					SheetID = worksheet.SheetID,
					Row = cellsStoreEnumerator.Row,
					Column = cellsStoreEnumerator.Column
				};
				if (cellsStoreEnumerator.Value is int)
				{
					formulaCell.Formula = worksheet._sharedFormulas[(int)cellsStoreEnumerator.Value].GetFormula(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, worksheet.Name);
				}
				else
				{
					formulaCell.Formula = cellsStoreEnumerator.Value.ToString();
				}
				if (!string.IsNullOrEmpty(formulaCell.Formula))
				{
					formulaCell.Tokens = lexer.Tokenize(formulaCell.Formula, Range.Worksheet.Name).ToList();
					worksheet._formulaTokens.SetValue(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, formulaCell.Tokens);
					depChain.Add(formulaCell);
					FollowChain(depChain, lexer, worksheet.Workbook, worksheet, formulaCell, options);
				}
			}
		}
	}

	private static void FollowChain(DependencyChain depChain, ILexer lexer, ExcelWorkbook wb, ExcelWorksheet ws, FormulaCell f, ExcelCalculationOption options)
	{
		Stack<FormulaCell> stack = new Stack<FormulaCell>();
		while (true)
		{
			if (f.tokenIx < f.Tokens.Count)
			{
				Token token = f.Tokens[f.tokenIx];
				if (token.TokenType == TokenType.ExcelAddress)
				{
					ExcelFormulaAddress excelFormulaAddress = new ExcelFormulaAddress(token.Value);
					if (excelFormulaAddress.Table != null)
					{
						excelFormulaAddress.SetRCFromTable(ws._package, new ExcelAddressBase(f.Row, f.Column, f.Row, f.Column));
					}
					if (excelFormulaAddress.WorkSheet == null && excelFormulaAddress.Collide(new ExcelAddressBase(f.Row, f.Column, f.Row, f.Column)) != 0 && !options.AllowCirculareReferences)
					{
						throw new CircularReferenceException($"Circular Reference in cell {ExcelCellBase.GetAddress(f.Row, f.Column)}");
					}
					if (excelFormulaAddress._fromRow > 0 && excelFormulaAddress._fromCol > 0)
					{
						if (string.IsNullOrEmpty(excelFormulaAddress.WorkSheet))
						{
							if (f.ws == null)
							{
								f.ws = ws;
							}
							else if (f.ws.SheetID != f.SheetID)
							{
								f.ws = wb.Worksheets.GetBySheetID(f.SheetID);
							}
						}
						else
						{
							f.ws = wb.Worksheets[excelFormulaAddress.WorkSheet];
						}
						if (f.ws != null)
						{
							f.iterator = new CellsStoreEnumerator<object>(f.ws._formulas, excelFormulaAddress.Start.Row, excelFormulaAddress.Start.Column, excelFormulaAddress.End.Row, excelFormulaAddress.End.Column);
							goto IL_06c3;
						}
					}
				}
				else if (token.TokenType == TokenType.NameValue)
				{
					ExcelAddressBase.SplitAddress(token.Value, out var _, out var ws2, out var address, (f.ws == null) ? "" : f.ws.Name);
					ExcelNamedRange excelNamedRange;
					if (!string.IsNullOrEmpty(ws2))
					{
						if (f.ws == null)
						{
							f.ws = wb.Worksheets[ws2];
						}
						excelNamedRange = (f.ws.Names.ContainsKey(token.Value) ? f.ws.Names[address] : ((!wb.Names.ContainsKey(address)) ? null : wb.Names[address]));
						if (excelNamedRange != null)
						{
							f.ws = excelNamedRange.Worksheet;
						}
					}
					else if (wb.Names.ContainsKey(address))
					{
						excelNamedRange = wb.Names[token.Value];
						if (string.IsNullOrEmpty(ws2))
						{
							f.ws = excelNamedRange.Worksheet;
						}
					}
					else
					{
						excelNamedRange = null;
					}
					if (excelNamedRange != null)
					{
						if (string.IsNullOrEmpty(excelNamedRange.NameFormula))
						{
							if (excelNamedRange.NameValue == null)
							{
								f.iterator = new CellsStoreEnumerator<object>(f.ws._formulas, excelNamedRange.Start.Row, excelNamedRange.Start.Column, excelNamedRange.End.Row, excelNamedRange.End.Column);
								goto IL_06c3;
							}
						}
						else
						{
							ulong cellID = ExcelCellBase.GetCellID(excelNamedRange.LocalSheetId, excelNamedRange.Index, 0);
							if (!depChain.index.ContainsKey(cellID))
							{
								FormulaCell formulaCell = new FormulaCell
								{
									SheetID = excelNamedRange.LocalSheetId,
									Row = excelNamedRange.Index,
									Column = 0
								};
								formulaCell.Formula = excelNamedRange.NameFormula;
								formulaCell.Tokens = ((excelNamedRange.LocalSheetId == -1) ? lexer.Tokenize(formulaCell.Formula).ToList() : lexer.Tokenize(formulaCell.Formula, wb.Worksheets.GetBySheetID(excelNamedRange.LocalSheetId).Name).ToList());
								depChain.Add(formulaCell);
								stack.Push(f);
								f = formulaCell;
								continue;
							}
							if (stack.Count > 0)
							{
								foreach (FormulaCell item in stack)
								{
									if (ExcelCellBase.GetCellID(item.SheetID, item.Row, item.Column) == cellID && !options.AllowCirculareReferences)
									{
										throw new CircularReferenceException($"Circular Reference in name {excelNamedRange.Name}");
									}
								}
							}
						}
					}
				}
				f.tokenIx++;
				continue;
			}
			depChain.CalcOrder.Add(f.Index);
			if (stack.Count > 0)
			{
				f = stack.Pop();
				goto IL_06c3;
			}
			break;
			IL_06c3:
			while (true)
			{
				if (f.iterator != null && f.iterator.Next())
				{
					object value = f.iterator.Value;
					if (value == null || value.ToString().Trim() == "")
					{
						continue;
					}
					ulong cellID2 = ExcelCellBase.GetCellID(f.ws.SheetID, f.iterator.Row, f.iterator.Column);
					if (!depChain.index.ContainsKey(cellID2))
					{
						FormulaCell formulaCell2 = new FormulaCell
						{
							SheetID = f.ws.SheetID,
							Row = f.iterator.Row,
							Column = f.iterator.Column
						};
						if (f.iterator.Value is int)
						{
							formulaCell2.Formula = f.ws._sharedFormulas[(int)value].GetFormula(f.iterator.Row, f.iterator.Column, ws.Name);
						}
						else
						{
							formulaCell2.Formula = value.ToString();
						}
						formulaCell2.ws = f.ws;
						formulaCell2.Tokens = lexer.Tokenize(formulaCell2.Formula, f.ws.Name).ToList();
						ws._formulaTokens.SetValue(formulaCell2.Row, formulaCell2.Column, formulaCell2.Tokens);
						depChain.Add(formulaCell2);
						stack.Push(f);
						f = formulaCell2;
						break;
					}
					if (stack.Count <= 0)
					{
						continue;
					}
					foreach (FormulaCell item2 in stack)
					{
						if (ExcelCellBase.GetCellID(item2.ws.SheetID, item2.iterator.Row, item2.iterator.Column) == cellID2)
						{
							if (!options.AllowCirculareReferences)
							{
								throw new CircularReferenceException($"Circular Reference in cell {item2.ws.Name}!{ExcelCellBase.GetAddress(f.Row, f.Column)}");
							}
							f = stack.Pop();
							break;
						}
					}
					continue;
				}
				f.tokenIx++;
				break;
			}
		}
	}
}
