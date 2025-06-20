using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing;

public abstract class ExcelDataProvider : IDisposable
{
	public interface IRangeInfo : IEnumerator<ICellInfo>, IDisposable, IEnumerator, IEnumerable<ICellInfo>, IEnumerable
	{
		bool IsEmpty { get; }

		bool IsMulti { get; }

		ExcelAddressBase Address { get; }

		ExcelWorksheet Worksheet { get; }

		int GetNCells();

		object GetValue(int row, int col);

		object GetOffset(int rowOffset, int colOffset);
	}

	public interface ICellInfo
	{
		string Address { get; }

		int Row { get; }

		int Column { get; }

		string Formula { get; }

		object Value { get; }

		double ValueDouble { get; }

		double ValueDoubleLogical { get; }

		bool IsHiddenRow { get; }

		bool IsExcelError { get; }

		IList<Token> Tokens { get; }
	}

	public interface INameInfo
	{
		ulong Id { get; set; }

		string Worksheet { get; set; }

		string Name { get; set; }

		string Formula { get; set; }

		IList<Token> Tokens { get; }

		object Value { get; set; }
	}

	public abstract int ExcelMaxColumns { get; }

	public abstract int ExcelMaxRows { get; }

	public abstract ExcelNamedRangeCollection GetWorksheetNames(string worksheet);

	public abstract ExcelNamedRangeCollection GetWorkbookNameValues();

	public abstract IRangeInfo GetRange(string worksheetName, int row, int column, string address);

	public abstract IRangeInfo GetRange(string worksheetName, string address);

	public abstract INameInfo GetName(string worksheet, string name);

	public abstract IEnumerable<object> GetRangeValues(string address);

	public abstract string GetRangeFormula(string worksheetName, int row, int column);

	public abstract List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column);

	public abstract bool IsRowHidden(string worksheetName, int row);

	public abstract object GetCellValue(string sheetName, int row, int col);

	public abstract ExcelCellAddress GetDimensionEnd(string worksheet);

	public abstract void Dispose();

	public abstract object GetRangeValue(string worksheetName, int row, int column);

	public abstract string GetFormat(object value, string format);

	public abstract void Reset();

	public abstract IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol);
}
