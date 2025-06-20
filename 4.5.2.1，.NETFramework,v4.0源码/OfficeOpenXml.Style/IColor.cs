using System.Drawing;

namespace OfficeOpenXml.Style;

internal interface IColor
{
	int Indexed { get; set; }

	string Rgb { get; }

	string Theme { get; }

	decimal Tint { get; set; }

	void SetColor(Color color);
}
