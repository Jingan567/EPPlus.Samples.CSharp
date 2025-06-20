using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing;

public sealed class ExcelDrawingFill : XmlHelper
{
	private string _fillPath;

	private XmlNode _fillNode;

	private eFillStyle _style;

	private XmlNode _fillTypeNode;

	private const string ColorPath = "/a:solidFill/a:srgbClr/@val";

	private const string alphaPath = "/a:solidFill/a:srgbClr/a:alpha/@val";

	public eFillStyle Style
	{
		get
		{
			if (_fillTypeNode == null)
			{
				return eFillStyle.SolidFill;
			}
			_style = GetStyleEnum(_fillTypeNode.Name);
			return _style;
		}
		set
		{
			if (value == eFillStyle.NoFill || value == eFillStyle.SolidFill)
			{
				_style = value;
				CreateFillTopNode(value);
				return;
			}
			throw new NotImplementedException("Fillstyle not implemented");
		}
	}

	public Color Color
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(_fillPath + "/a:solidFill/a:srgbClr/@val");
			if (xmlNodeString == "")
			{
				return Color.FromArgb(79, 129, 189);
			}
			return Color.FromArgb(int.Parse(xmlNodeString, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			if (_fillTypeNode == null)
			{
				_style = eFillStyle.SolidFill;
			}
			else if (_style != eFillStyle.SolidFill)
			{
				throw new Exception("FillStyle must be set to SolidFill");
			}
			CreateNode(_fillPath, insertFirst: false);
			SetXmlNodeString(_fillPath + "/a:solidFill/a:srgbClr/@val", value.ToArgb().ToString("X8").Substring(2));
		}
	}

	public int Transparancy
	{
		get
		{
			return 100 - GetXmlNodeInt(_fillPath + "/a:solidFill/a:srgbClr/a:alpha/@val") / 1000;
		}
		set
		{
			if (_fillTypeNode == null)
			{
				_style = eFillStyle.SolidFill;
				Color = Color.FromArgb(79, 129, 189);
			}
			else if (_style != eFillStyle.SolidFill)
			{
				throw new Exception("FillStyle must be set to SolidFill");
			}
			SetXmlNodeString(_fillPath + "/a:solidFill/a:srgbClr/a:alpha/@val", ((100 - value) * 1000).ToString());
		}
	}

	internal ExcelDrawingFill(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath)
		: base(nameSpaceManager, topNode)
	{
		_fillPath = fillPath;
		_fillNode = topNode.SelectSingleNode(_fillPath, base.NameSpaceManager);
		base.SchemaNodeOrder = new string[16]
		{
			"tickLblPos", "spPr", "txPr", "dLblPos", "crossAx", "printSettings", "showVal", "prstGeom", "noFill", "solidFill",
			"blipFill", "gradFill", "noFill", "pattFill", "ln", "prstDash"
		};
		if (_fillNode != null)
		{
			_fillTypeNode = topNode.SelectSingleNode("solidFill");
			if (_fillTypeNode == null)
			{
				_fillTypeNode = topNode.SelectSingleNode("noFill");
			}
			if (_fillTypeNode == null)
			{
				_fillTypeNode = topNode.SelectSingleNode("blipFill");
			}
			if (_fillTypeNode == null)
			{
				_fillTypeNode = topNode.SelectSingleNode("gradFill");
			}
			if (_fillTypeNode == null)
			{
				_fillTypeNode = topNode.SelectSingleNode("pattFill");
			}
		}
	}

	private void CreateFillTopNode(eFillStyle value)
	{
		if (_fillTypeNode != null)
		{
			base.TopNode.RemoveChild(_fillTypeNode);
		}
		CreateNode(_fillPath + "/a:" + GetStyleText(value), insertFirst: false);
		_fillNode = base.TopNode.SelectSingleNode(_fillPath + "/a:" + GetStyleText(value), base.NameSpaceManager);
	}

	private eFillStyle GetStyleEnum(string name)
	{
		return name switch
		{
			"noFill" => eFillStyle.NoFill, 
			"blipFill" => eFillStyle.BlipFill, 
			"gradFill" => eFillStyle.GradientFill, 
			"grpFill" => eFillStyle.GroupFill, 
			"pattFill" => eFillStyle.PatternFill, 
			_ => eFillStyle.SolidFill, 
		};
	}

	private string GetStyleText(eFillStyle style)
	{
		return style switch
		{
			eFillStyle.BlipFill => "blipFill", 
			eFillStyle.GradientFill => "gradFill", 
			eFillStyle.GroupFill => "grpFill", 
			eFillStyle.NoFill => "noFill", 
			eFillStyle.PatternFill => "pattFill", 
			_ => "solidFill", 
		};
	}
}
