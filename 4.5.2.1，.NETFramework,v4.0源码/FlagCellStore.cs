using OfficeOpenXml;

internal class FlagCellStore : CellStore<byte>
{
	internal void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
	{
		CellFlags value2 = (CellFlags)GetValue(Row, Col);
		if (value)
		{
			SetValue(Row, Col, (byte)(value2 | cellFlags));
		}
		else
		{
			SetValue(Row, Col, (byte)(value2 & ~cellFlags));
		}
	}

	internal bool GetFlagValue(int Row, int Col, CellFlags cellFlags)
	{
		return ((byte)cellFlags & GetValue(Row, Col)) != 0;
	}
}
