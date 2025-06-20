using System;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public class ExcelRow : IRangeID
{
	private ExcelWorksheet _worksheet;

	private XmlElement _rowElement;

	internal string _styleName = "";

	[Obsolete]
	public ulong RowID => GetRowID(_worksheet.SheetID, Row);

	internal XmlNode Node => _rowElement;

	public bool Hidden
	{
		get
		{
			return ((RowInternal)_worksheet.GetValueInner(Row, 0))?.Hidden ?? false;
		}
		set
		{
			GetRowInternal().Hidden = value;
		}
	}

	public double Height
	{
		get
		{
			RowInternal rowInternal = (RowInternal)_worksheet.GetValueInner(Row, 0);
			if (rowInternal == null || rowInternal.Height < 0.0)
			{
				return _worksheet.DefaultRowHeight;
			}
			return rowInternal.Height;
		}
		set
		{
			RowInternal rowInternal = GetRowInternal();
			if (_worksheet._package.DoAdjustDrawings)
			{
				int[,] drawingHeight = _worksheet.Drawings.GetDrawingHeight();
				rowInternal.Height = value;
				_worksheet.Drawings.AdjustHeight(drawingHeight);
			}
			else
			{
				rowInternal.Height = value;
			}
			if (rowInternal.Hidden && value != 0.0)
			{
				Hidden = false;
			}
			rowInternal.CustomHeight = true;
		}
	}

	public bool CustomHeight
	{
		get
		{
			return ((RowInternal)_worksheet.GetValueInner(Row, 0))?.CustomHeight ?? false;
		}
		set
		{
			GetRowInternal().CustomHeight = value;
		}
	}

	public string StyleName
	{
		get
		{
			return _styleName;
		}
		set
		{
			StyleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
			_styleName = value;
		}
	}

	public int StyleID
	{
		get
		{
			return _worksheet.GetStyleInner(Row, 0);
		}
		set
		{
			_worksheet.SetStyleInner(Row, 0, value);
		}
	}

	public int Row { get; set; }

	public bool Collapsed
	{
		get
		{
			return ((RowInternal)_worksheet.GetValueInner(Row, 0))?.Collapsed ?? false;
		}
		set
		{
			GetRowInternal().Collapsed = value;
		}
	}

	public int OutlineLevel
	{
		get
		{
			return ((RowInternal)_worksheet.GetValueInner(Row, 0))?.OutlineLevel ?? 0;
		}
		set
		{
			GetRowInternal().OutlineLevel = (short)value;
		}
	}

	public bool Phonetic
	{
		get
		{
			return ((RowInternal)_worksheet.GetValueInner(Row, 0))?.Phonetic ?? false;
		}
		set
		{
			GetRowInternal().Phonetic = value;
		}
	}

	public ExcelStyle Style => _worksheet.Workbook.Styles.GetStyleObject(StyleID, _worksheet.PositionID, Row + ":" + Row);

	public bool PageBreak
	{
		get
		{
			return ((RowInternal)_worksheet.GetValueInner(Row, 0))?.PageBreak ?? false;
		}
		set
		{
			GetRowInternal().PageBreak = value;
		}
	}

	public bool Merged
	{
		get
		{
			return _worksheet.MergedCells[Row, 0] != null;
		}
		set
		{
			_worksheet.MergedCells.Add(new ExcelAddressBase(Row, 1, Row, 16384), doValidate: true);
		}
	}

	[Obsolete]
	ulong IRangeID.RangeID
	{
		get
		{
			return RowID;
		}
		set
		{
			Row = (int)(value >> 29);
		}
	}

	internal ExcelRow(ExcelWorksheet Worksheet, int row)
	{
		_worksheet = Worksheet;
		Row = row;
	}

	private RowInternal GetRowInternal()
	{
		RowInternal rowInternal = (RowInternal)_worksheet.GetValueInner(Row, 0);
		if (rowInternal == null)
		{
			rowInternal = new RowInternal();
			_worksheet.SetValueInner(Row, 0, rowInternal);
		}
		return rowInternal;
	}

	internal static ulong GetRowID(int sheetID, int row)
	{
		return (ulong)(sheetID + ((long)row << 29));
	}

	internal void Clone(ExcelWorksheet added)
	{
		ExcelRow excelRow = added.Row(Row);
		excelRow.Collapsed = Collapsed;
		excelRow.Height = Height;
		excelRow.CustomHeight = excelRow.CustomHeight;
		excelRow.Hidden = Hidden;
		excelRow.OutlineLevel = OutlineLevel;
		excelRow.PageBreak = PageBreak;
		excelRow.Phonetic = Phonetic;
		excelRow._styleName = _styleName;
		excelRow.StyleID = StyleID;
	}
}
