using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public abstract class ExcelCellBase
{
	private delegate string dlgTransl(string part, int row, int col, int rowIncr, int colIncr);

	internal static void SplitCellID(ulong cellID, out int sheet, out int row, out int col)
	{
		sheet = (int)(cellID % 32768);
		col = (int)(cellID >> 15) & 0x3FF;
		row = (int)(cellID >> 29);
	}

	internal static ulong GetCellID(int SheetID, int row, int col)
	{
		return (ulong)(SheetID + ((long)col << 15) + ((long)row << 29));
	}

	public static string TranslateFromR1C1(string value, int row, int col)
	{
		return Translate(value, ToAbs, row, col, -1, -1);
	}

	public static string TranslateToR1C1(string value, int row, int col)
	{
		return Translate(value, ToR1C1, row, col, -1, -1);
	}

	private static string Translate(string value, dlgTransl addressTranslator, int row, int col, int rowIncr, int colIncr)
	{
		if (value == "")
		{
			return "";
		}
		IEnumerable<Token> enumerable = new Lexer(SourceCodeTokenizer.R1C1, new SyntacticAnalyzer()).Tokenize(value, null);
		foreach (Token item in enumerable)
		{
			if (item.TokenType == TokenType.ExcelAddress || item.TokenType.Equals(TokenType.NameValue) || item.TokenType == TokenType.ExcelAddressR1C1)
			{
				string value2 = addressTranslator(item.Value, row, col, rowIncr, colIncr);
				item.Value = value2;
			}
		}
		return string.Join("", enumerable.Select((Token x) => x.Value).ToArray());
	}

	private static string ToR1C1(string part, int row, int col, int rowIncr, int colIncr)
	{
		int num = part.IndexOf('!');
		string text = "";
		if (num > 0)
		{
			text = part.Substring(0, num + 1);
			part = part.Substring(num + 1);
		}
		int num2 = part.IndexOf(':');
		if (num2 > 0)
		{
			string text2 = ToR1C1_1(part.Substring(0, num2), row, col, rowIncr, colIncr);
			string text3 = ToR1C1_1(part.Substring(num2 + 1), row, col, rowIncr, colIncr);
			if (text2.Equals(text3))
			{
				return text2;
			}
			return text + text2 + ":" + text3;
		}
		return text + ToR1C1_1(part, row, col, rowIncr, colIncr);
	}

	private static string ToR1C1_1(string part, int row, int col, int rowIncr, int colIncr)
	{
		StringBuilder stringBuilder = new StringBuilder();
		if (GetRowCol(part, out var row2, out var col2, throwException: false, out var fixedRow, out var fixedCol))
		{
			if (row2 == 0 && col2 == 0)
			{
				return part;
			}
			if (row2 > 0)
			{
				stringBuilder.Append(fixedRow ? $"R{row2}" : ((row2 == row) ? "R" : $"R[{row2 - row}]"));
			}
			if (col2 > 0)
			{
				stringBuilder.Append(fixedCol ? $"C{col2}" : ((col2 == col) ? "C" : $"C[{col2 - col}]"));
			}
			return stringBuilder.ToString();
		}
		return part;
	}

	private static string ToAbs(string part, int row, int col, int rowIncr, int colIncr)
	{
		int num = part.IndexOf('!');
		string text = "";
		if (num > 0)
		{
			text = part.Substring(0, num + 1);
			part = part.Substring(num + 1);
		}
		int num2 = part.IndexOf(':');
		if (num2 > 0)
		{
			string text2 = ToAbs_1(part.Substring(0, num2), row, col, rowIncr, colIncr);
			string text3 = ToAbs_1(part.Substring(num2 + 1), row, col, rowIncr, colIncr);
			if (text2.Equals(text3))
			{
				return text2;
			}
			return text + text2 + ":" + text3;
		}
		return text + ToAbs_1(part, row, col, rowIncr, colIncr);
	}

	private static string ToAbs_1(string part, int row, int col, int rowIncr, int colIncr)
	{
		string text = ConvertUtil._invariantTextInfo.ToUpper(part);
		int num = text.IndexOf("R");
		int num2 = text.IndexOf("C");
		if (num != 0 && num2 != 0)
		{
			return part;
		}
		if (part.Length == 1)
		{
			if (num == 0)
			{
				return $"{row}:{row}";
			}
			string columnLetter = GetColumnLetter(col);
			return $"{columnLetter}:{columnLetter}";
		}
		bool fixedAddr;
		if (num2 == -1)
		{
			int rC = GetRC(part.Substring(1), row, out fixedAddr);
			if (rC > int.MinValue)
			{
				return GetAddressRow(rC, fixedAddr);
			}
			return part;
		}
		bool fixedAddr2;
		if (num == -1)
		{
			int rC2 = GetRC(part.Substring(1), col, out fixedAddr2);
			if (rC2 > int.MinValue)
			{
				return GetAddressCol(rC2, fixedAddr2);
			}
			return part;
		}
		int num3;
		if (1 == num2)
		{
			num3 = row;
			fixedAddr = false;
		}
		else
		{
			num3 = GetRC(part.Substring(1, num2 - 1), row, out fixedAddr);
		}
		int num4;
		if (part.Length - 1 == num2)
		{
			num4 = col;
			fixedAddr2 = false;
		}
		else
		{
			num4 = GetRC(part.Substring(num2 + 1, part.Length - num2 - 1), col, out fixedAddr2);
		}
		if (num3 > int.MinValue && num4 > int.MinValue)
		{
			return GetAddress(num3, fixedAddr, num4, fixedAddr2);
		}
		return part;
	}

	private static string AddToRowColumnTranslator(string Address, int row, int col, int rowIncr, int colIncr)
	{
		if (Address == "#REF!")
		{
			return Address;
		}
		if (GetRowCol(Address, out var row2, out var col2, throwException: false))
		{
			if (row2 == 0 || col2 == 0)
			{
				return Address;
			}
			if (rowIncr != 0 && row != 0 && row2 >= row && Address.IndexOf('$', 1) == -1)
			{
				if (row2 < row - rowIncr)
				{
					return "#REF!";
				}
				row2 += rowIncr;
			}
			if (colIncr != 0 && col != 0 && col2 >= col && !ConvertUtil._invariantCompareInfo.IsPrefix(Address, "$"))
			{
				if (col2 < col - colIncr)
				{
					return "#REF!";
				}
				col2 += colIncr;
			}
			Address = GetAddress(row2, Address.IndexOf('$', 1) > -1, col2, ConvertUtil._invariantCompareInfo.IsPrefix(Address, "$"));
		}
		return Address;
	}

	private static string GetRCFmt(int v)
	{
		if (v >= 0)
		{
			if (v <= 0)
			{
				return "";
			}
			return v.ToString();
		}
		return $"[{v}]";
	}

	private static int GetRC(string value, int OffsetValue, out bool fixedAddr)
	{
		if (value == "")
		{
			fixedAddr = false;
			return OffsetValue;
		}
		int result;
		if (value[0] == '[' && value[value.Length - 1] == ']')
		{
			fixedAddr = false;
			if (int.TryParse(value.Substring(1, value.Length - 2), NumberStyles.Any, CultureInfo.InvariantCulture, out result))
			{
				return OffsetValue + result;
			}
			return int.MinValue;
		}
		fixedAddr = true;
		if (int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
		{
			return result;
		}
		return int.MinValue;
	}

	protected internal static string GetColumnLetter(int iColumnNumber)
	{
		return GetColumnLetter(iColumnNumber, fixedCol: false);
	}

	protected internal static string GetColumnLetter(int iColumnNumber, bool fixedCol)
	{
		if (iColumnNumber < 1)
		{
			return "#REF!";
		}
		string text = "";
		do
		{
			text = (char)(65 + (iColumnNumber - 1) % 26) + text;
			iColumnNumber = (iColumnNumber - (iColumnNumber - 1) % 26) / 26;
		}
		while (iColumnNumber > 0);
		if (!fixedCol)
		{
			return text;
		}
		return "$" + text;
	}

	internal static bool GetRowColFromAddress(string CellAddress, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn)
	{
		bool fixedFromRow;
		bool fixedFromColumn;
		bool fixedToRow;
		bool fixedToColumn;
		return GetRowColFromAddress(CellAddress, out FromRow, out FromColumn, out ToRow, out ToColumn, out fixedFromRow, out fixedFromColumn, out fixedToRow, out fixedToColumn);
	}

	internal static bool GetRowColFromAddress(string CellAddress, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn, out bool fixedFromRow, out bool fixedFromColumn, out bool fixedToRow, out bool fixedToColumn)
	{
		if (CellAddress.IndexOf('[') > 0)
		{
			FromRow = -1;
			FromColumn = -1;
			ToRow = -1;
			ToColumn = -1;
			fixedFromRow = false;
			fixedFromColumn = false;
			fixedToRow = false;
			fixedToColumn = false;
			return false;
		}
		CellAddress = ConvertUtil._invariantTextInfo.ToUpper(CellAddress);
		if (CellAddress.IndexOf(' ') > 0)
		{
			CellAddress = CellAddress.Substring(0, CellAddress.IndexOf(' '));
		}
		bool rowColFromAddress;
		if (CellAddress.IndexOf(':') < 0)
		{
			rowColFromAddress = GetRowColFromAddress(CellAddress, out FromRow, out FromColumn, out fixedFromRow, out fixedFromColumn);
			ToColumn = FromColumn;
			ToRow = FromRow;
			fixedToRow = fixedFromRow;
			fixedToColumn = fixedFromColumn;
		}
		else
		{
			string[] array = CellAddress.Split(':');
			rowColFromAddress = GetRowColFromAddress(array[0], out FromRow, out FromColumn, out fixedFromRow, out fixedFromColumn);
			if (rowColFromAddress)
			{
				rowColFromAddress = GetRowColFromAddress(array[1], out ToRow, out ToColumn, out fixedToRow, out fixedToColumn);
			}
			else
			{
				GetRowColFromAddress(array[1], out ToRow, out ToColumn, out fixedToRow, out fixedToColumn);
			}
			if (FromColumn <= 0)
			{
				FromColumn = 1;
			}
			if (FromRow <= 0)
			{
				FromRow = 1;
			}
			if (ToColumn <= 0)
			{
				ToColumn = 16384;
			}
			if (ToRow <= 0)
			{
				ToRow = 1048576;
			}
		}
		return rowColFromAddress;
	}

	internal static bool GetRowColFromAddress(string CellAddress, out int Row, out int Column)
	{
		return GetRowCol(CellAddress, out Row, out Column, throwException: true);
	}

	internal static bool GetRowColFromAddress(string CellAddress, out int row, out int col, out bool fixedRow, out bool fixedCol)
	{
		return GetRowCol(CellAddress, out row, out col, throwException: true, out fixedRow, out fixedCol);
	}

	internal static bool IsAlpha(char c)
	{
		if (c >= 'A')
		{
			return c <= 'Z';
		}
		return false;
	}

	internal static bool GetRowCol(string address, out int row, out int col, bool throwException)
	{
		bool fixedRow;
		bool fixedCol;
		return GetRowCol(address, out row, out col, throwException, out fixedRow, out fixedCol);
	}

	internal static bool GetRowCol(string address, out int row, out int col, bool throwException, out bool fixedRow, out bool fixedCol)
	{
		bool flag = true;
		int num = 0;
		int num2 = 0;
		col = 0;
		row = 0;
		fixedRow = false;
		fixedCol = false;
		if (ConvertUtil._invariantCompareInfo.IsSuffix(address, "#REF!"))
		{
			row = 0;
			col = 0;
			return true;
		}
		int num3 = address.IndexOf('!');
		if (num3 >= 0)
		{
			num = num3 + 1;
		}
		address = ConvertUtil._invariantTextInfo.ToUpper(address);
		for (int i = num; i < address.Length; i++)
		{
			char c = address[i];
			if (flag && c >= 'A' && c <= 'Z' && num2 <= 3)
			{
				col *= 26;
				col += c - 64;
				num2++;
				continue;
			}
			if (c >= '0' && c <= '9')
			{
				row *= 10;
				row += c - 48;
				flag = false;
				continue;
			}
			if (c == '$')
			{
				if (IsAlpha(address[i + 1]))
				{
					num++;
					fixedCol = true;
				}
				else
				{
					flag = false;
					fixedRow = true;
				}
				continue;
			}
			row = 0;
			col = 0;
			if (throwException)
			{
				throw new Exception($"Invalid Address format {address}");
			}
			return false;
		}
		if (row == 0)
		{
			return col != 0;
		}
		return true;
	}

	private static int GetColumn(string sCol)
	{
		int num = 0;
		int num2 = sCol.Length - 1;
		for (int num3 = num2; num3 >= 0; num3--)
		{
			num += (sCol[num3] - 64) * (int)Math.Pow(26.0, num2 - num3);
		}
		return num;
	}

	public static string GetAddressRow(int Row, bool Absolute = false)
	{
		if (Absolute)
		{
			return $"${Row}:${Row}";
		}
		return $"{Row}:{Row}";
	}

	public static string GetAddressCol(int Col, bool Absolute = false)
	{
		string columnLetter = GetColumnLetter(Col);
		if (Absolute)
		{
			return $"${columnLetter}:${columnLetter}";
		}
		return $"{columnLetter}:{columnLetter}";
	}

	public static string GetAddress(int Row, int Column)
	{
		return GetAddress(Row, Column, Absolute: false);
	}

	public static string GetAddress(int Row, bool AbsoluteRow, int Column, bool AbsoluteCol)
	{
		if (Row < 1 || Row > 1048576 || Column < 1 || Column > 16384)
		{
			return "#REF!";
		}
		return (AbsoluteCol ? "$" : "") + GetColumnLetter(Column) + (AbsoluteRow ? "$" : "") + Row;
	}

	public static string GetAddress(int Row, int Column, bool Absolute)
	{
		if (Row == 0 || Column == 0)
		{
			return "#REF!";
		}
		if (Absolute)
		{
			return "$" + GetColumnLetter(Column) + "$" + Row;
		}
		return GetColumnLetter(Column) + Row;
	}

	public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn)
	{
		return GetAddress(FromRow, FromColumn, ToRow, ToColumn, Absolute: false);
	}

	public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool Absolute)
	{
		if (FromRow == ToRow && FromColumn == ToColumn)
		{
			return GetAddress(FromRow, FromColumn, Absolute);
		}
		if (FromRow == 1 && ToRow == 1048576)
		{
			string text = (Absolute ? "$" : "");
			return text + GetColumnLetter(FromColumn) + ":" + text + GetColumnLetter(ToColumn);
		}
		if (FromColumn == 1 && ToColumn == 16384)
		{
			string text2 = (Absolute ? "$" : "");
			return text2 + FromRow + ":" + text2 + ToRow;
		}
		return GetAddress(FromRow, FromColumn, Absolute) + ":" + GetAddress(ToRow, ToColumn, Absolute);
	}

	public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool FixedFromRow, bool FixedFromColumn, bool FixedToRow, bool FixedToColumn)
	{
		if (FromRow == ToRow && FromColumn == ToColumn)
		{
			return GetAddress(FromRow, FixedFromRow, FromColumn, FixedFromColumn);
		}
		if (FromRow == 1 && ToRow == 1048576)
		{
			return GetColumnLetter(FromColumn, FixedFromColumn) + ":" + GetColumnLetter(ToColumn, FixedToColumn);
		}
		if (FromColumn == 1 && ToColumn == 16384)
		{
			return (FixedFromRow ? "$" : "") + FromRow + ":" + (FixedToRow ? "$" : "") + ToRow;
		}
		return GetAddress(FromRow, FixedFromRow, FromColumn, FixedFromColumn) + ":" + GetAddress(ToRow, FixedToRow, ToColumn, FixedToColumn);
	}

	public static string GetFullAddress(string worksheetName, string address)
	{
		return GetFullAddress(worksheetName, address, fullRowCol: true);
	}

	internal static string GetFullAddress(string worksheetName, string address, bool fullRowCol)
	{
		if (address.IndexOf("!") == -1 || address == "#REF!")
		{
			if (fullRowCol)
			{
				string[] array = address.Split(':');
				if (array.Length != 0)
				{
					address = $"'{worksheetName}'!{array[0]}";
					if (array.Length > 1)
					{
						address += $":{array[1]}";
					}
				}
			}
			else
			{
				ExcelAddressBase excelAddressBase = new ExcelAddressBase(address);
				address = (((excelAddressBase._fromRow != 1 || excelAddressBase._toRow != 1048576) && (excelAddressBase._fromCol != 1 || excelAddressBase._toCol != 16384)) ? GetFullAddress(worksheetName, address, fullRowCol: true) : $"'{worksheetName}'!{GetColumnLetter(excelAddressBase._fromCol)}{excelAddressBase._fromRow}:{GetColumnLetter(excelAddressBase._toCol)}{excelAddressBase._toRow}");
			}
		}
		return address;
	}

	public static bool IsValidAddress(string address)
	{
		if (string.IsNullOrEmpty(address.Trim()))
		{
			return false;
		}
		address = ConvertUtil._invariantTextInfo.ToUpper(address);
		string[] array = address.Split(',');
		foreach (string text in array)
		{
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			bool flag = false;
			for (int j = 0; j < text.Length; j++)
			{
				if (IsCol(text[j]))
				{
					if (!flag)
					{
						if (text2 != "")
						{
							return false;
						}
						text3 += text[j];
						if (text3.Length > 3)
						{
							return false;
						}
					}
					else
					{
						if (text4 != "")
						{
							return false;
						}
						text5 += text[j];
						if (text5.Length > 3)
						{
							return false;
						}
					}
				}
				else if (IsRow(text[j]))
				{
					if (!flag)
					{
						text2 += text[j];
						if (text2.Length > 7)
						{
							return false;
						}
					}
					else
					{
						text4 += text[j];
						if (text4.Length > 7)
						{
							return false;
						}
					}
				}
				else if (text[j] == ':')
				{
					if (flag || j == text.Length - 1)
					{
						return false;
					}
					flag = true;
				}
				else
				{
					if (text[j] != '$')
					{
						return false;
					}
					if (j == text.Length - 1 || text[j + 1] == ':' || (j > 1 && IsCol(text[j - 1]) && IsCol(text[j + 1])) || (j > 1 && IsRow(text[j - 1]) && IsRow(text[j + 1])))
					{
						return false;
					}
				}
			}
			bool flag2;
			if (text2 != "" && text3 != "" && text4 == "" && text5 == "")
			{
				flag2 = GetColumn(text3) <= 16384 && int.Parse(text2) <= 1048576;
			}
			else if (text2 != "" && text4 != "" && text3 != "" && text5 != "")
			{
				int num = int.Parse(text4);
				int column = GetColumn(text5);
				flag2 = GetColumn(text3) <= column && int.Parse(text2) <= num && column <= 16384 && num <= 1048576;
			}
			else if (text2 == "" && text4 == "" && text3 != "" && text5 != "")
			{
				int column2 = GetColumn(text5);
				flag2 = GetColumn(text3) <= column2 && column2 <= 16384;
			}
			else
			{
				if (!(text2 != "") || !(text4 != "") || !(text3 == "") || !(text5 == ""))
				{
					return false;
				}
				int num2 = int.Parse(text4);
				flag2 = int.Parse(text2) <= num2 && num2 <= 1048576;
			}
			if (!flag2)
			{
				return false;
			}
		}
		return true;
	}

	private static bool IsCol(char c)
	{
		if (c >= 'A')
		{
			return c <= 'Z';
		}
		return false;
	}

	private static bool IsRow(char r)
	{
		if (r >= '0')
		{
			return r <= '9';
		}
		return false;
	}

	public static bool IsValidCellAddress(string cellAddress)
	{
		bool result = false;
		try
		{
			if (GetRowColFromAddress(cellAddress, out var Row, out var Column))
			{
				result = ((Row > 0 && Column > 0 && Row <= 1048576 && Column <= 16384) ? true : false);
			}
		}
		catch
		{
		}
		return result;
	}

	internal static string UpdateFormulaReferences(string formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, string currentSheet, string modifiedSheet, bool setFixed = false)
	{
		new Dictionary<string, object>();
		try
		{
			IEnumerable<Token> enumerable = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty).Tokenize(formula);
			string text = "";
			foreach (Token item in enumerable)
			{
				if (item.TokenType == TokenType.ExcelAddress)
				{
					ExcelAddressBase excelAddressBase = new ExcelAddressBase(item.Value);
					bool flag = (string.IsNullOrEmpty(excelAddressBase._ws) && currentSheet.Equals(modifiedSheet, StringComparison.CurrentCultureIgnoreCase)) || modifiedSheet.Equals(excelAddressBase._ws, StringComparison.CurrentCultureIgnoreCase);
					if (!setFixed && (!string.IsNullOrEmpty(excelAddressBase._wb) || !flag))
					{
						text += excelAddressBase.Address;
						continue;
					}
					if (!string.IsNullOrEmpty(excelAddressBase._ws))
					{
						text += $"'{excelAddressBase._ws}'!";
					}
					if (rowIncrement > 0)
					{
						excelAddressBase = excelAddressBase.AddRow(afterRow, rowIncrement, setFixed);
					}
					else if (rowIncrement < 0)
					{
						excelAddressBase = excelAddressBase.DeleteRow(afterRow, -rowIncrement, setFixed);
					}
					if (colIncrement > 0)
					{
						excelAddressBase = excelAddressBase.AddColumn(afterColumn, colIncrement, setFixed);
					}
					else if (colIncrement < 0)
					{
						excelAddressBase = excelAddressBase.DeleteColumn(afterColumn, -colIncrement, setFixed);
					}
					if (excelAddressBase == null || !excelAddressBase.IsValidRowCol())
					{
						text += "#REF!";
						continue;
					}
					string[] array = excelAddressBase.Address.Split('!');
					text = ((array.Length <= 1) ? (text + excelAddressBase.Address) : (text + array[1]));
				}
				else
				{
					text += item.Value;
				}
			}
			return text;
		}
		catch
		{
			return formula;
		}
	}

	internal static string UpdateFormulaSheetReferences(string formula, string oldSheetName, string newSheetName)
	{
		if (string.IsNullOrEmpty(oldSheetName))
		{
			throw new ArgumentNullException("oldSheetName");
		}
		if (string.IsNullOrEmpty(newSheetName))
		{
			throw new ArgumentNullException("newSheetName");
		}
		new Dictionary<string, object>();
		try
		{
			IEnumerable<Token> enumerable = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty).Tokenize(formula);
			string text = "";
			foreach (Token item in enumerable)
			{
				if (item.TokenType == TokenType.ExcelAddress)
				{
					ExcelAddressBase excelAddressBase = new ExcelAddressBase(item.Value);
					if (excelAddressBase == null || !excelAddressBase.IsValidRowCol())
					{
						text += "#REF!";
						continue;
					}
					excelAddressBase.ChangeWorksheet(oldSheetName, newSheetName);
					text += excelAddressBase.Address;
				}
				else
				{
					text += item.Value;
				}
			}
			return text;
		}
		catch
		{
			return formula;
		}
	}
}
