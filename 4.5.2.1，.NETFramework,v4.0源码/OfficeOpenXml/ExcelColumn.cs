using System;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public class ExcelColumn : IRangeID
{
	private ExcelWorksheet _worksheet;

	private XmlElement _colElement;

	internal int _columnMin;

	internal int _columnMax;

	internal bool _hidden;

	internal double _width;

	internal string _styleName = "";

	public int ColumnMin => _columnMin;

	public int ColumnMax
	{
		get
		{
			return _columnMax;
		}
		set
		{
			if (value < _columnMin && value > 16384)
			{
				throw new Exception("ColumnMax out of range");
			}
			CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, 0, 0, 16384);
			while (cellsStoreEnumerator.Next())
			{
				ExcelColumn excelColumn = cellsStoreEnumerator.Value._value as ExcelColumn;
				if (cellsStoreEnumerator.Column > _columnMin && excelColumn.ColumnMax <= value && cellsStoreEnumerator.Column != _columnMin)
				{
					throw new Exception($"ColumnMax can not span over existing column {excelColumn.ColumnMin}.");
				}
			}
			_columnMax = value;
		}
	}

	internal ulong ColumnID => GetColumnID(_worksheet.SheetID, ColumnMin);

	public bool Hidden
	{
		get
		{
			return _hidden;
		}
		set
		{
			if (_worksheet._package.DoAdjustDrawings)
			{
				int[,] drawingWidths = _worksheet.Drawings.GetDrawingWidths();
				_hidden = value;
				_worksheet.Drawings.AdjustWidth(drawingWidths);
			}
			else
			{
				_hidden = value;
			}
		}
	}

	internal double VisualWidth
	{
		get
		{
			if (_hidden || (Collapsed && OutlineLevel > 0))
			{
				return 0.0;
			}
			return _width;
		}
	}

	public double Width
	{
		get
		{
			return _width;
		}
		set
		{
			if (_worksheet._package.DoAdjustDrawings)
			{
				int[,] drawingWidths = _worksheet.Drawings.GetDrawingWidths();
				_width = value;
				_worksheet.Drawings.AdjustWidth(drawingWidths);
			}
			else
			{
				_width = value;
			}
			if (_hidden && value != 0.0)
			{
				_hidden = false;
			}
		}
	}

	public bool BestFit { get; set; }

	public bool Collapsed { get; set; }

	public int OutlineLevel { get; set; }

	public bool Phonetic { get; set; }

	public ExcelStyle Style
	{
		get
		{
			string columnLetter = ExcelCellBase.GetColumnLetter(ColumnMin);
			string columnLetter2 = ExcelCellBase.GetColumnLetter(ColumnMax);
			return _worksheet.Workbook.Styles.GetStyleObject(StyleID, _worksheet.PositionID, columnLetter + ":" + columnLetter2);
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
			return _worksheet.GetStyleInner(0, ColumnMin);
		}
		set
		{
			_worksheet.SetStyleInner(0, ColumnMin, value);
		}
	}

	public bool PageBreak { get; set; }

	public bool Merged
	{
		get
		{
			return _worksheet.MergedCells[ColumnMin, 0] != null;
		}
		set
		{
			_worksheet.MergedCells.Add(new ExcelAddressBase(1, ColumnMin, 1048576, ColumnMax), doValidate: true);
		}
	}

	ulong IRangeID.RangeID
	{
		get
		{
			return ColumnID;
		}
		set
		{
			int columnMin = _columnMin;
			_columnMin = (int)(value >> 15) & 0x3FF;
			_columnMax += columnMin - ColumnMin;
			if (_columnMax > 16384)
			{
				_columnMax = 16384;
			}
		}
	}

	protected internal ExcelColumn(ExcelWorksheet Worksheet, int col)
	{
		_worksheet = Worksheet;
		_columnMin = col;
		_columnMax = col;
		_width = _worksheet.DefaultColWidth;
	}

	public override string ToString()
	{
		return string.Format("Column Range: {0} to {1}", _colElement.GetAttribute("min"), _colElement.GetAttribute("min"));
	}

	public void AutoFit()
	{
		_worksheet.Cells[1, _columnMin, 1048576, _columnMax].AutoFitColumns();
	}

	public void AutoFit(double MinimumWidth)
	{
		_worksheet.Cells[1, _columnMin, 1048576, _columnMax].AutoFitColumns(MinimumWidth);
	}

	public void AutoFit(double MinimumWidth, double MaximumWidth)
	{
		_worksheet.Cells[1, _columnMin, 1048576, _columnMax].AutoFitColumns(MinimumWidth, MaximumWidth);
	}

	internal static ulong GetColumnID(int sheetID, int column)
	{
		return (ulong)(sheetID + ((long)column << 15));
	}

	internal ExcelColumn Clone(ExcelWorksheet added)
	{
		return Clone(added, ColumnMin);
	}

	internal ExcelColumn Clone(ExcelWorksheet added, int col)
	{
		ExcelColumn excelColumn = added.Column(col);
		excelColumn.ColumnMax = ColumnMax;
		excelColumn.BestFit = BestFit;
		excelColumn.Collapsed = Collapsed;
		excelColumn.OutlineLevel = OutlineLevel;
		excelColumn.PageBreak = PageBreak;
		excelColumn.Phonetic = Phonetic;
		excelColumn._styleName = _styleName;
		excelColumn.StyleID = StyleID;
		excelColumn.Width = Width;
		excelColumn.Hidden = Hidden;
		return excelColumn;
	}
}
