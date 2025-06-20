using System;
using System.Xml;

namespace OfficeOpenXml;

public class ExcelWorksheetView : XmlHelper
{
	public class ExcelWorksheetPanes : XmlHelper
	{
		private XmlElement _selectionNode;

		private const string _activeCellPath = "@activeCell";

		private const string _selectionRangePath = "@sqref";

		public string ActiveCell
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@activeCell");
				if (xmlNodeString == "")
				{
					return "A1";
				}
				return xmlNodeString;
			}
			set
			{
				if (_selectionNode == null)
				{
					CreateSelectionElement();
				}
				ExcelCellBase.GetRowColFromAddress(value, out int FromRow, out int FromColumn, out int _, out int _);
				SetXmlNodeString("@activeCell", value);
				if (((XmlElement)base.TopNode).GetAttribute("sqref") == "")
				{
					SelectedRange = ExcelCellBase.GetAddress(FromRow, FromColumn);
				}
			}
		}

		public string SelectedRange
		{
			get
			{
				string xmlNodeString = GetXmlNodeString("@sqref");
				if (xmlNodeString == "")
				{
					return "A1";
				}
				return xmlNodeString;
			}
			set
			{
				if (_selectionNode == null)
				{
					CreateSelectionElement();
				}
				ExcelCellBase.GetRowColFromAddress(value, out int FromRow, out int FromColumn, out int _, out int _);
				SetXmlNodeString("@sqref", value);
				if (((XmlElement)base.TopNode).GetAttribute("activeCell") == "")
				{
					ActiveCell = ExcelCellBase.GetAddress(FromRow, FromColumn);
				}
			}
		}

		internal ExcelWorksheetPanes(XmlNamespaceManager ns, XmlNode topNode)
			: base(ns, topNode)
		{
			if (topNode.Name == "selection")
			{
				_selectionNode = topNode as XmlElement;
			}
		}

		private void CreateSelectionElement()
		{
			_selectionNode = base.TopNode.OwnerDocument.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			base.TopNode.AppendChild(_selectionNode);
			base.TopNode = _selectionNode;
		}
	}

	private ExcelWorksheet _worksheet;

	private XmlElement _selectionNode;

	private string _paneNodePath = "d:pane";

	private string _selectionNodePath = "d:selection";

	protected internal XmlElement SheetViewElement => (XmlElement)base.TopNode;

	private XmlElement SelectionNode
	{
		get
		{
			_selectionNode = SheetViewElement.SelectSingleNode("//d:selection", _worksheet.NameSpaceManager) as XmlElement;
			if (_selectionNode == null)
			{
				_selectionNode = _worksheet.WorksheetXml.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				SheetViewElement.AppendChild(_selectionNode);
			}
			return _selectionNode;
		}
	}

	public string ActiveCell
	{
		get
		{
			return Panes[Panes.GetUpperBound(0)].ActiveCell;
		}
		set
		{
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(value);
			if (!excelAddressBase.IsSingleCell)
			{
				throw new InvalidOperationException("ActiveCell must be a single cell.");
			}
			ExcelAddressBase sd = new ExcelAddressBase(SelectedRange.Replace(" ", ","));
			Panes[Panes.GetUpperBound(0)].ActiveCell = value;
			if (!IsActiveCellInSelection(excelAddressBase, sd))
			{
				SelectedRange = value;
			}
		}
	}

	public string SelectedRange
	{
		get
		{
			return Panes[Panes.GetUpperBound(0)].SelectedRange;
		}
		set
		{
			ExcelAddressBase ac = new ExcelAddressBase(ActiveCell);
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(value.Replace(" ", ","));
			Panes[Panes.GetUpperBound(0)].SelectedRange = value;
			if (!IsActiveCellInSelection(ac, excelAddressBase))
			{
				ActiveCell = new ExcelCellAddress(excelAddressBase._fromRow, excelAddressBase._fromCol).Address;
			}
		}
	}

	public bool TabSelected
	{
		get
		{
			return GetXmlNodeBool("@tabSelected");
		}
		set
		{
			SetTabSelected(value);
		}
	}

	public bool TabSelectedMulti
	{
		get
		{
			return GetXmlNodeBool("@tabSelected");
		}
		set
		{
			SetTabSelected(value, allowMultiple: true);
		}
	}

	public bool PageLayoutView
	{
		get
		{
			return GetXmlNodeString("@view") == "pageLayout";
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("@view", "pageLayout");
			}
			else
			{
				SheetViewElement.RemoveAttribute("view");
			}
		}
	}

	public bool PageBreakView
	{
		get
		{
			return GetXmlNodeString("@view") == "pageBreakPreview";
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("@view", "pageBreakPreview");
			}
			else
			{
				SheetViewElement.RemoveAttribute("view");
			}
		}
	}

	public bool ShowGridLines
	{
		get
		{
			return GetXmlNodeBool("@showGridLines");
		}
		set
		{
			SetXmlNodeString("@showGridLines", value ? "1" : "0");
		}
	}

	public bool ShowHeaders
	{
		get
		{
			return GetXmlNodeBool("@showRowColHeaders");
		}
		set
		{
			SetXmlNodeString("@showRowColHeaders", value ? "1" : "0");
		}
	}

	public int ZoomScale
	{
		get
		{
			return GetXmlNodeInt("@zoomScale");
		}
		set
		{
			if (value < 10 || value > 400)
			{
				throw new ArgumentOutOfRangeException("Zoome scale out of range (10-400)");
			}
			SetXmlNodeString("@zoomScale", value.ToString());
		}
	}

	public bool RightToLeft
	{
		get
		{
			return GetXmlNodeBool("@rightToLeft");
		}
		set
		{
			SetXmlNodeString("@rightToLeft", value ? "1" : "0");
		}
	}

	internal bool WindowProtection
	{
		get
		{
			return GetXmlNodeBool("@windowProtection", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@windowProtection", value, removeIf: false);
		}
	}

	public ExcelWorksheetPanes[] Panes { get; internal set; }

	internal ExcelWorksheetView(XmlNamespaceManager ns, XmlNode node, ExcelWorksheet xlWorksheet)
		: base(ns, node)
	{
		_worksheet = xlWorksheet;
		base.SchemaNodeOrder = new string[4] { "sheetViews", "sheetView", "pane", "selection" };
		Panes = LoadPanes();
	}

	private ExcelWorksheetPanes[] LoadPanes()
	{
		XmlNodeList xmlNodeList = base.TopNode.SelectNodes("//d:selection", base.NameSpaceManager);
		if (xmlNodeList.Count == 0)
		{
			return new ExcelWorksheetPanes[1]
			{
				new ExcelWorksheetPanes(base.NameSpaceManager, base.TopNode)
			};
		}
		ExcelWorksheetPanes[] array = new ExcelWorksheetPanes[xmlNodeList.Count];
		int num = 0;
		foreach (XmlElement item in xmlNodeList)
		{
			array[num++] = new ExcelWorksheetPanes(base.NameSpaceManager, item);
		}
		return array;
	}

	private bool IsActiveCellInSelection(ExcelAddressBase ac, ExcelAddressBase sd)
	{
		ExcelAddressBase.eAddressCollition eAddressCollition = sd.Collide(ac);
		if (eAddressCollition == ExcelAddressBase.eAddressCollition.Equal || eAddressCollition == ExcelAddressBase.eAddressCollition.Inside)
		{
			return true;
		}
		if (sd.Addresses != null)
		{
			foreach (ExcelAddress address in sd.Addresses)
			{
				eAddressCollition = address.Collide(ac);
				if (eAddressCollition == ExcelAddressBase.eAddressCollition.Equal || eAddressCollition == ExcelAddressBase.eAddressCollition.Inside)
				{
					return true;
				}
			}
		}
		return false;
	}

	public void SetTabSelected(bool isSelected = true, bool allowMultiple = false)
	{
		if (isSelected)
		{
			SheetViewElement.SetAttribute("tabSelected", "1");
			if (!allowMultiple)
			{
				foreach (ExcelWorksheet worksheet in _worksheet._package.Workbook.Worksheets)
				{
					worksheet.View.TabSelected = false;
				}
			}
			if (_worksheet.Workbook.WorkbookXml.SelectSingleNode("//d:workbookView", _worksheet.NameSpaceManager) is XmlElement xmlElement)
			{
				xmlElement.SetAttribute("activeTab", (_worksheet.PositionID - 1).ToString());
			}
		}
		else
		{
			SetXmlNodeString("@tabSelected", "0");
		}
	}

	public void FreezePanes(int Row, int Column)
	{
		if (Row == 1 && Column == 1)
		{
			UnFreezePanes();
		}
		string selectedRange = SelectedRange;
		string activeCell = ActiveCell;
		XmlElement xmlElement = base.TopNode.SelectSingleNode(_paneNodePath, base.NameSpaceManager) as XmlElement;
		if (xmlElement == null)
		{
			CreateNode(_paneNodePath);
			xmlElement = base.TopNode.SelectSingleNode(_paneNodePath, base.NameSpaceManager) as XmlElement;
		}
		xmlElement.RemoveAll();
		if (Column > 1)
		{
			xmlElement.SetAttribute("xSplit", (Column - 1).ToString());
		}
		if (Row > 1)
		{
			xmlElement.SetAttribute("ySplit", (Row - 1).ToString());
		}
		xmlElement.SetAttribute("topLeftCell", ExcelCellBase.GetAddress(Row, Column));
		xmlElement.SetAttribute("state", "frozen");
		RemoveSelection();
		if (Row > 1 && Column == 1)
		{
			xmlElement.SetAttribute("activePane", "bottomLeft");
			XmlElement xmlElement2 = base.TopNode.OwnerDocument.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement2.SetAttribute("pane", "bottomLeft");
			if (activeCell != "")
			{
				xmlElement2.SetAttribute("activeCell", activeCell);
			}
			if (selectedRange != "")
			{
				xmlElement2.SetAttribute("sqref", selectedRange);
			}
			xmlElement2.SetAttribute("sqref", selectedRange);
			base.TopNode.InsertAfter(xmlElement2, xmlElement);
		}
		else if (Column > 1 && Row == 1)
		{
			xmlElement.SetAttribute("activePane", "topRight");
			XmlElement xmlElement3 = base.TopNode.OwnerDocument.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement3.SetAttribute("pane", "topRight");
			if (activeCell != "")
			{
				xmlElement3.SetAttribute("activeCell", activeCell);
			}
			if (selectedRange != "")
			{
				xmlElement3.SetAttribute("sqref", selectedRange);
			}
			base.TopNode.InsertAfter(xmlElement3, xmlElement);
		}
		else
		{
			xmlElement.SetAttribute("activePane", "bottomRight");
			XmlElement xmlElement4 = base.TopNode.OwnerDocument.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement4.SetAttribute("pane", "topRight");
			string address = ExcelCellBase.GetAddress(1, Column);
			xmlElement4.SetAttribute("activeCell", address);
			xmlElement4.SetAttribute("sqref", address);
			xmlElement.ParentNode.InsertAfter(xmlElement4, xmlElement);
			XmlElement xmlElement5 = base.TopNode.OwnerDocument.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			address = ExcelCellBase.GetAddress(Row, 1);
			xmlElement5.SetAttribute("pane", "bottomLeft");
			xmlElement5.SetAttribute("activeCell", address);
			xmlElement5.SetAttribute("sqref", address);
			xmlElement4.ParentNode.InsertAfter(xmlElement5, xmlElement4);
			XmlElement xmlElement6 = base.TopNode.OwnerDocument.CreateElement("selection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement6.SetAttribute("pane", "bottomRight");
			if (activeCell != "")
			{
				xmlElement6.SetAttribute("activeCell", activeCell);
			}
			if (selectedRange != "")
			{
				xmlElement6.SetAttribute("sqref", selectedRange);
			}
			xmlElement5.ParentNode.InsertAfter(xmlElement6, xmlElement5);
		}
		Panes = LoadPanes();
	}

	private void RemoveSelection()
	{
		foreach (XmlNode item in base.TopNode.SelectNodes(_selectionNodePath, base.NameSpaceManager))
		{
			item.ParentNode.RemoveChild(item);
		}
	}

	public void UnFreezePanes()
	{
		string selectedRange = SelectedRange;
		string activeCell = ActiveCell;
		if (base.TopNode.SelectSingleNode(_paneNodePath, base.NameSpaceManager) is XmlElement xmlElement)
		{
			xmlElement.ParentNode.RemoveChild(xmlElement);
		}
		RemoveSelection();
		Panes = LoadPanes();
		SelectedRange = selectedRange;
		ActiveCell = activeCell;
	}
}
