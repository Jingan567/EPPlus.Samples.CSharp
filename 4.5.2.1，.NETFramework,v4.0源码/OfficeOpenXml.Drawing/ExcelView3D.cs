using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing;

public sealed class ExcelView3D : XmlHelper
{
	private const string perspectivePath = "c:perspective/@val";

	private const string rotXPath = "c:rotX/@val";

	private const string rotYPath = "c:rotY/@val";

	private const string rAngAxPath = "c:rAngAx/@val";

	private const string depthPercentPath = "c:depthPercent/@val";

	private const string heightPercentPath = "c:hPercent/@val";

	public decimal Perspective
	{
		get
		{
			return GetXmlNodeInt("c:perspective/@val");
		}
		set
		{
			SetXmlNodeString("c:perspective/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal RotX
	{
		get
		{
			return GetXmlNodeDecimal("c:rotX/@val");
		}
		set
		{
			CreateNode("c:rotX/@val");
			SetXmlNodeString("c:rotX/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal RotY
	{
		get
		{
			return GetXmlNodeDecimal("c:rotY/@val");
		}
		set
		{
			CreateNode("c:rotY/@val");
			SetXmlNodeString("c:rotY/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public bool RightAngleAxes
	{
		get
		{
			return GetXmlNodeBool("c:rAngAx/@val");
		}
		set
		{
			SetXmlNodeBool("c:rAngAx/@val", value);
		}
	}

	public int DepthPercent
	{
		get
		{
			return GetXmlNodeInt("c:depthPercent/@val");
		}
		set
		{
			if (value < 0 || value > 2000)
			{
				throw new ArgumentOutOfRangeException("Value must be between 0 and 2000");
			}
			SetXmlNodeString("c:depthPercent/@val", value.ToString());
		}
	}

	public int HeightPercent
	{
		get
		{
			return GetXmlNodeInt("c:hPercent/@val");
		}
		set
		{
			if (value < 5 || value > 500)
			{
				throw new ArgumentOutOfRangeException("Value must be between 5 and 500");
			}
			SetXmlNodeString("c:hPercent/@val", value.ToString());
		}
	}

	internal ExcelView3D(XmlNamespaceManager ns, XmlNode node)
		: base(ns, node)
	{
		base.SchemaNodeOrder = new string[6] { "rotX", "hPercent", "rotY", "depthPercent", "rAngAx", "perspective" };
	}
}
