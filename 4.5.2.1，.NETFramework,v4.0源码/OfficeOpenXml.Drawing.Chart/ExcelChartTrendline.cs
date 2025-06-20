using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartTrendline : XmlHelper
{
	private const string TRENDLINEPATH = "c:trendlineType/@val";

	private const string NAMEPATH = "c:name";

	private const string ORDERPATH = "c:order/@val";

	private const string PERIODPATH = "c:period/@val";

	private const string FORWARDPATH = "c:forward/@val";

	private const string BACKWARDPATH = "c:backward/@val";

	private const string INTERCEPTPATH = "c:intercept/@val";

	private const string DISPLAYRSQUAREDVALUEPATH = "c:dispRSqr/@val";

	private const string DISPLAYEQUATIONPATH = "c:dispEq/@val";

	public eTrendLine Type
	{
		get
		{
			return GetXmlNodeString("c:trendlineType/@val").ToLower(CultureInfo.InvariantCulture) switch
			{
				"exp" => eTrendLine.Exponential, 
				"log" => eTrendLine.Logarithmic, 
				"poly" => eTrendLine.Polynomial, 
				"movingavg" => eTrendLine.MovingAvgerage, 
				"power" => eTrendLine.Power, 
				_ => eTrendLine.Linear, 
			};
		}
		set
		{
			switch (value)
			{
			case eTrendLine.Exponential:
				SetXmlNodeString("c:trendlineType/@val", "exp");
				break;
			case eTrendLine.Logarithmic:
				SetXmlNodeString("c:trendlineType/@val", "log");
				break;
			case eTrendLine.Polynomial:
				SetXmlNodeString("c:trendlineType/@val", "poly");
				Order = 2m;
				break;
			case eTrendLine.MovingAvgerage:
				SetXmlNodeString("c:trendlineType/@val", "movingAvg");
				Period = 2m;
				break;
			case eTrendLine.Power:
				SetXmlNodeString("c:trendlineType/@val", "power");
				break;
			default:
				SetXmlNodeString("c:trendlineType/@val", "linear");
				break;
			}
		}
	}

	public string Name
	{
		get
		{
			return GetXmlNodeString("c:name");
		}
		set
		{
			SetXmlNodeString("c:name", value, removeIfBlank: true);
		}
	}

	public decimal Order
	{
		get
		{
			return GetXmlNodeDecimal("c:order/@val");
		}
		set
		{
			if (Type == eTrendLine.MovingAvgerage)
			{
				throw new ArgumentException("Can't set period for trendline type MovingAvgerage");
			}
			DeleteAllNode("c:period/@val");
			SetXmlNodeString("c:order/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal Period
	{
		get
		{
			return GetXmlNodeDecimal("c:period/@val");
		}
		set
		{
			if (Type == eTrendLine.Polynomial)
			{
				throw new ArgumentException("Can't set period for trendline type Polynomial");
			}
			DeleteAllNode("c:order/@val");
			SetXmlNodeString("c:period/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal Forward
	{
		get
		{
			return GetXmlNodeDecimal("c:forward/@val");
		}
		set
		{
			SetXmlNodeString("c:forward/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal Backward
	{
		get
		{
			return GetXmlNodeDecimal("c:backward/@val");
		}
		set
		{
			SetXmlNodeString("c:backward/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal Intercept
	{
		get
		{
			return GetXmlNodeDecimal("c:intercept/@val");
		}
		set
		{
			SetXmlNodeString("c:intercept/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public bool DisplayRSquaredValue
	{
		get
		{
			return GetXmlNodeBool("c:dispRSqr/@val", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("c:dispRSqr/@val", value, removeIf: true);
		}
	}

	public bool DisplayEquation
	{
		get
		{
			return GetXmlNodeBool("c:dispEq/@val", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("c:dispEq/@val", value, removeIf: true);
		}
	}

	internal ExcelChartTrendline(XmlNamespaceManager namespaceManager, XmlNode topNode)
		: base(namespaceManager, topNode)
	{
		base.SchemaNodeOrder = new string[10] { "name", "trendlineType", "order", "period", "forward", "backward", "intercept", "dispRSqr", "dispEq", "trendlineLbl" };
	}
}
