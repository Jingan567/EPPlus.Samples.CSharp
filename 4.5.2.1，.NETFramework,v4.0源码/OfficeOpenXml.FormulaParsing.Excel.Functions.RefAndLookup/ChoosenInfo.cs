using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class ChoosenInfo : ExcelDataProvider.IRangeInfo, IEnumerator<ExcelDataProvider.ICellInfo>, IDisposable, IEnumerator, IEnumerable<ExcelDataProvider.ICellInfo>, IEnumerable
{
	private string[] chosenIndeces;

	public bool IsEmpty => false;

	public bool IsMulti => true;

	public ExcelAddressBase Address => null;

	public ExcelWorksheet Worksheet => null;

	public ExcelDataProvider.ICellInfo Current => null;

	object IEnumerator.Current => chosenIndeces[0];

	public ChoosenInfo(string[] chosenIndeces)
	{
		this.chosenIndeces = chosenIndeces;
	}

	public int GetNCells()
	{
		return 0;
	}

	public object GetValue(int row, int col)
	{
		return null;
	}

	public object GetOffset(int rowOffset, int colOffset)
	{
		return null;
	}

	public void Dispose()
	{
	}

	public bool MoveNext()
	{
		throw new NotImplementedException();
	}

	public void Reset()
	{
		throw new NotImplementedException();
	}

	public IEnumerator<ExcelDataProvider.ICellInfo> GetEnumerator()
	{
		throw new NotImplementedException();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		throw new NotImplementedException();
	}
}
