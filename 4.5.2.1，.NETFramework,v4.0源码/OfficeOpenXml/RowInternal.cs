namespace OfficeOpenXml;

internal class RowInternal
{
	internal double Height = -1.0;

	internal bool Hidden;

	internal bool Collapsed;

	internal short OutlineLevel;

	internal bool PageBreak;

	internal bool Phonetic;

	internal bool CustomHeight;

	internal int MergeID;

	internal RowInternal Clone()
	{
		return new RowInternal
		{
			Height = Height,
			Hidden = Hidden,
			Collapsed = Collapsed,
			OutlineLevel = OutlineLevel,
			PageBreak = PageBreak,
			Phonetic = Phonetic,
			CustomHeight = CustomHeight,
			MergeID = MergeID
		};
	}
}
