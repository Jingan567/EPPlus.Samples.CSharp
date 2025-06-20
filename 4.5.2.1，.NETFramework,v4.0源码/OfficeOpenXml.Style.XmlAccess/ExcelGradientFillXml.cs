using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelGradientFillXml : ExcelFillXml
{
	private const string _typePath = "d:gradientFill/@type";

	private const string _degreePath = "d:gradientFill/@degree";

	private const string _gradientColor1Path = "d:gradientFill/d:stop[@position=\"0\"]/d:color";

	private const string _gradientColor2Path = "d:gradientFill/d:stop[@position=\"1\"]/d:color";

	private const string _bottomPath = "d:gradientFill/@bottom";

	private const string _topPath = "d:gradientFill/@top";

	private const string _leftPath = "d:gradientFill/@left";

	private const string _rightPath = "d:gradientFill/@right";

	public ExcelFillGradientType Type { get; internal set; }

	public double Degree { get; internal set; }

	public ExcelColorXml GradientColor1 { get; private set; }

	public ExcelColorXml GradientColor2 { get; private set; }

	public double Bottom { get; internal set; }

	public double Top { get; internal set; }

	public double Left { get; internal set; }

	public double Right { get; internal set; }

	internal override string Id => string.Concat(base.Id, Degree.ToString(), GradientColor1.Id, GradientColor2.Id, Type, Left.ToString(), Right.ToString(), Bottom.ToString(), Top.ToString());

	internal ExcelGradientFillXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
		GradientColor1 = new ExcelColorXml(nameSpaceManager);
		GradientColor2 = new ExcelColorXml(nameSpaceManager);
	}

	internal ExcelGradientFillXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		Degree = GetXmlNodeDouble("d:gradientFill/@degree");
		Type = ((!(GetXmlNodeString("d:gradientFill/@type") == "path")) ? ExcelFillGradientType.Linear : ExcelFillGradientType.Path);
		GradientColor1 = new ExcelColorXml(nsm, topNode.SelectSingleNode("d:gradientFill/d:stop[@position=\"0\"]/d:color", nsm));
		GradientColor2 = new ExcelColorXml(nsm, topNode.SelectSingleNode("d:gradientFill/d:stop[@position=\"1\"]/d:color", nsm));
		Top = GetXmlNodeDouble("d:gradientFill/@top");
		Bottom = GetXmlNodeDouble("d:gradientFill/@bottom");
		Left = GetXmlNodeDouble("d:gradientFill/@left");
		Right = GetXmlNodeDouble("d:gradientFill/@right");
	}

	internal override ExcelFillXml Copy()
	{
		return new ExcelGradientFillXml(base.NameSpaceManager)
		{
			PatternType = _fillPatternType,
			BackgroundColor = _backgroundColor.Copy(),
			PatternColor = _patternColor.Copy(),
			GradientColor1 = GradientColor1.Copy(),
			GradientColor2 = GradientColor2.Copy(),
			Type = Type,
			Degree = Degree,
			Top = Top,
			Bottom = Bottom,
			Left = Left,
			Right = Right
		};
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		CreateNode("d:gradientFill");
		if (Type == ExcelFillGradientType.Path)
		{
			SetXmlNodeString("d:gradientFill/@type", "path");
		}
		if (!double.IsNaN(Degree))
		{
			SetXmlNodeString("d:gradientFill/@degree", Degree.ToString(CultureInfo.InvariantCulture));
		}
		if (GradientColor1 != null)
		{
			XmlNode xmlNode = base.TopNode.SelectSingleNode("d:gradientFill", base.NameSpaceManager);
			XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("stop", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.SetAttribute("position", "0");
			xmlNode.AppendChild(xmlElement);
			XmlElement xmlElement2 = xmlNode.OwnerDocument.CreateElement("color", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.AppendChild(xmlElement2);
			GradientColor1.CreateXmlNode(xmlElement2);
			xmlElement = xmlNode.OwnerDocument.CreateElement("stop", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.SetAttribute("position", "1");
			xmlNode.AppendChild(xmlElement);
			xmlElement2 = xmlNode.OwnerDocument.CreateElement("color", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.AppendChild(xmlElement2);
			GradientColor2.CreateXmlNode(xmlElement2);
		}
		if (!double.IsNaN(Top))
		{
			SetXmlNodeString("d:gradientFill/@top", Top.ToString("F5", CultureInfo.InvariantCulture));
		}
		if (!double.IsNaN(Bottom))
		{
			SetXmlNodeString("d:gradientFill/@bottom", Bottom.ToString("F5", CultureInfo.InvariantCulture));
		}
		if (!double.IsNaN(Left))
		{
			SetXmlNodeString("d:gradientFill/@left", Left.ToString("F5", CultureInfo.InvariantCulture));
		}
		if (!double.IsNaN(Right))
		{
			SetXmlNodeString("d:gradientFill/@right", Right.ToString("F5", CultureInfo.InvariantCulture));
		}
		return topNode;
	}
}
