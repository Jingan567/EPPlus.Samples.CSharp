namespace OfficeOpenXml;

public class FontSizeInfo
{
	public float Height { get; set; }

	public float Width { get; set; }

	public FontSizeInfo(float height, float width)
	{
		Width = width;
		Height = height;
	}
}
