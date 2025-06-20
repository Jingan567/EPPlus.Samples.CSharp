using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelBubbleChartSerie : ExcelChartSerie
{
	private ExcelChartSerieDataLabel _DataLabel;

	private const string BUBBLE3D_PATH = "c:bubble3D/@val";

	private const string INVERTIFNEGATIVE_PATH = "c:invertIfNegative/@val";

	private const string BUBBLESIZE_TOPPATH = "c:bubbleSize";

	private const string BUBBLESIZE_PATH = "c:bubbleSize/c:numRef/c:f";

	public ExcelChartSerieDataLabel DataLabel
	{
		get
		{
			if (_DataLabel == null)
			{
				_DataLabel = new ExcelChartSerieDataLabel(_ns, _node);
			}
			return _DataLabel;
		}
	}

	internal bool Bubble3D
	{
		get
		{
			return GetXmlNodeBool("c:bubble3D/@val", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("c:bubble3D/@val", value);
		}
	}

	internal bool InvertIfNegative
	{
		get
		{
			return GetXmlNodeBool("c:invertIfNegative/@val", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("c:invertIfNegative/@val", value);
		}
	}

	public override string Series
	{
		get
		{
			return base.Series;
		}
		set
		{
			base.Series = value;
			if (string.IsNullOrEmpty(BubbleSize))
			{
				GenerateLit();
			}
		}
	}

	public string BubbleSize
	{
		get
		{
			return GetXmlNodeString("c:bubbleSize/c:numRef/c:f");
		}
		set
		{
			if (string.IsNullOrEmpty(value))
			{
				GenerateLit();
				return;
			}
			SetXmlNodeString("c:bubbleSize/c:numRef/c:f", ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));
			XmlNode xmlNode = base.TopNode.SelectSingleNode(string.Format("{0}/c:numCache", "c:bubbleSize/c:numRef/c:f"), _ns);
			xmlNode?.ParentNode.RemoveChild(xmlNode);
			DeleteNode(string.Format("{0}/c:numLit", "c:bubbleSize"));
		}
	}

	internal ExcelBubbleChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
	}

	internal void GenerateLit()
	{
		ExcelAddress excelAddress = new ExcelAddress(Series);
		int num = 0;
		StringBuilder stringBuilder = new StringBuilder();
		for (int i = excelAddress._fromRow; i <= excelAddress._toRow; i++)
		{
			for (int j = excelAddress._fromCol; j <= excelAddress._toCol; j++)
			{
				stringBuilder.AppendFormat("<c:pt idx=\"{0}\"><c:v>1</c:v></c:pt>", num++);
			}
		}
		CreateNode("c:bubbleSize/c:numLit", insertFirst: true);
		base.TopNode.SelectSingleNode(string.Format("{0}/c:numLit", "c:bubbleSize"), _ns).InnerXml = $"<c:formatCode>General</c:formatCode><c:ptCount val=\"{num}\"/>{stringBuilder.ToString()}";
	}
}
