using System;
using System.Drawing;

namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfColor : DxfStyleBase<ExcelDxfColor>
{
	public int? Theme { get; set; }

	public int? Index { get; set; }

	public bool? Auto { get; set; }

	public double? Tint { get; set; }

	public Color? Color { get; set; }

	protected internal override string Id => GetAsString(Theme) + "|" + GetAsString(Index) + "|" + GetAsString(Auto) + "|" + GetAsString(Tint) + "|" + GetAsString((!Color.HasValue) ? "" : Color.Value.ToArgb().ToString("x"));

	protected internal override bool HasValue
	{
		get
		{
			if (!Theme.HasValue && !Index.HasValue && !Auto.HasValue && !Tint.HasValue)
			{
				return Color.HasValue;
			}
			return true;
		}
	}

	public ExcelDxfColor(ExcelStyles styles)
		: base(styles)
	{
	}

	protected internal override ExcelDxfColor Clone()
	{
		return new ExcelDxfColor(_styles)
		{
			Theme = Theme,
			Index = Index,
			Color = Color,
			Auto = Auto,
			Tint = Tint
		};
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		throw new NotImplementedException();
	}
}
