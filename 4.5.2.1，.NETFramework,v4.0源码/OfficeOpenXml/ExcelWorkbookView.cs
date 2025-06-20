using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public class ExcelWorkbookView : XmlHelper
{
	private const string LEFT_PATH = "d:bookViews/d:workbookView/@xWindow";

	private const string TOP_PATH = "d:bookViews/d:workbookView/@yWindow";

	private const string WIDTH_PATH = "d:bookViews/d:workbookView/@windowWidth";

	private const string HEIGHT_PATH = "d:bookViews/d:workbookView/@windowHeight";

	private const string MINIMIZED_PATH = "d:bookViews/d:workbookView/@minimized";

	private const string SHOWVERTICALSCROLL_PATH = "d:bookViews/d:workbookView/@showVerticalScroll";

	private const string SHOWHORIZONTALSCR_PATH = "d:bookViews/d:workbookView/@showHorizontalScroll";

	private const string SHOWSHEETTABS_PATH = "d:bookViews/d:workbookView/@showSheetTabs";

	private const string ACTIVETAB_PATH = "d:bookViews/d:workbookView/@activeTab";

	public int Left
	{
		get
		{
			return GetXmlNodeInt("d:bookViews/d:workbookView/@xWindow");
		}
		internal set
		{
			SetXmlNodeString("d:bookViews/d:workbookView/@xWindow", value.ToString());
		}
	}

	public int Top
	{
		get
		{
			return GetXmlNodeInt("d:bookViews/d:workbookView/@yWindow");
		}
		internal set
		{
			SetXmlNodeString("d:bookViews/d:workbookView/@yWindow", value.ToString());
		}
	}

	public int Width
	{
		get
		{
			return GetXmlNodeInt("d:bookViews/d:workbookView/@windowWidth");
		}
		internal set
		{
			SetXmlNodeString("d:bookViews/d:workbookView/@windowWidth", value.ToString());
		}
	}

	public int Height
	{
		get
		{
			return GetXmlNodeInt("d:bookViews/d:workbookView/@windowHeight");
		}
		internal set
		{
			SetXmlNodeString("d:bookViews/d:workbookView/@windowHeight", value.ToString());
		}
	}

	public bool Minimized
	{
		get
		{
			return GetXmlNodeBool("d:bookViews/d:workbookView/@minimized");
		}
		set
		{
			SetXmlNodeString("d:bookViews/d:workbookView/@minimized", value.ToString());
		}
	}

	public bool ShowVerticalScrollBar
	{
		get
		{
			return GetXmlNodeBool("d:bookViews/d:workbookView/@showVerticalScroll", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:bookViews/d:workbookView/@showVerticalScroll", value, removeIf: true);
		}
	}

	public bool ShowHorizontalScrollBar
	{
		get
		{
			return GetXmlNodeBool("d:bookViews/d:workbookView/@showHorizontalScroll", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:bookViews/d:workbookView/@showHorizontalScroll", value, removeIf: true);
		}
	}

	public bool ShowSheetTabs
	{
		get
		{
			return GetXmlNodeBool("d:bookViews/d:workbookView/@showSheetTabs", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:bookViews/d:workbookView/@showSheetTabs", value, removeIf: true);
		}
	}

	public int ActiveTab
	{
		get
		{
			int xmlNodeInt = GetXmlNodeInt("d:bookViews/d:workbookView/@activeTab");
			if (xmlNodeInt < 0)
			{
				return 0;
			}
			return xmlNodeInt;
		}
		set
		{
			SetXmlNodeString("d:bookViews/d:workbookView/@activeTab", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelWorkbookView(XmlNamespaceManager ns, XmlNode node, ExcelWorkbook wb)
		: base(ns, node)
	{
		base.SchemaNodeOrder = wb.SchemaNodeOrder;
	}

	public void SetWindowSize(int left, int top, int width, int height)
	{
		Left = left;
		Top = top;
		Width = width;
		Height = height;
	}
}
