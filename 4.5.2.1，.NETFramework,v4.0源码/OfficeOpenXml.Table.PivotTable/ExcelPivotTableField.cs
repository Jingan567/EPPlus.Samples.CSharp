using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableField : XmlHelper
{
	internal ExcelPivotTable _table;

	internal ExcelPivotTablePageFieldSettings _pageFieldSettings;

	private ExcelPivotTableFieldGroup _grouping;

	internal XmlHelperInstance _cacheFieldHelper;

	internal ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> _items;

	public int Index { get; set; }

	internal int BaseIndex { get; set; }

	public string Name
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@name");
			if (xmlNodeString == "")
			{
				return _cacheFieldHelper.GetXmlNodeString("@name");
			}
			return xmlNodeString;
		}
		set
		{
			SetXmlNodeString("@name", value);
		}
	}

	public bool Compact
	{
		get
		{
			return GetXmlNodeBool("@compact");
		}
		set
		{
			SetXmlNodeBool("@compact", value);
		}
	}

	public bool Outline
	{
		get
		{
			return GetXmlNodeBool("@outline");
		}
		set
		{
			SetXmlNodeBool("@outline", value);
		}
	}

	public bool SubtotalTop
	{
		get
		{
			return GetXmlNodeBool("@subtotalTop");
		}
		set
		{
			SetXmlNodeBool("@subtotalTop", value);
		}
	}

	public bool MultipleItemSelectionAllowed
	{
		get
		{
			return GetXmlNodeBool("@multipleItemSelectionAllowed");
		}
		set
		{
			SetXmlNodeBool("@multipleItemSelectionAllowed", value);
		}
	}

	public bool ShowAll
	{
		get
		{
			return GetXmlNodeBool("@showAll");
		}
		set
		{
			SetXmlNodeBool("@showAll", value);
		}
	}

	public bool ShowDropDowns
	{
		get
		{
			return GetXmlNodeBool("@showDropDowns");
		}
		set
		{
			SetXmlNodeBool("@showDropDowns", value);
		}
	}

	public bool ShowInFieldList
	{
		get
		{
			return GetXmlNodeBool("@showInFieldList");
		}
		set
		{
			SetXmlNodeBool("@showInFieldList", value);
		}
	}

	public bool ShowAsCaption
	{
		get
		{
			return GetXmlNodeBool("@showPropAsCaption");
		}
		set
		{
			SetXmlNodeBool("@showPropAsCaption", value);
		}
	}

	public bool ShowMemberPropertyInCell
	{
		get
		{
			return GetXmlNodeBool("@showPropCell");
		}
		set
		{
			SetXmlNodeBool("@showPropCell", value);
		}
	}

	public bool ShowMemberPropertyToolTip
	{
		get
		{
			return GetXmlNodeBool("@showPropTip");
		}
		set
		{
			SetXmlNodeBool("@showPropTip", value);
		}
	}

	public eSortType Sort
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@sortType");
			if (!(xmlNodeString == ""))
			{
				return (eSortType)Enum.Parse(typeof(eSortType), xmlNodeString, ignoreCase: true);
			}
			return eSortType.None;
		}
		set
		{
			if (value == eSortType.None)
			{
				DeleteNode("@sortType");
			}
			else
			{
				SetXmlNodeString("@sortType", value.ToString().ToLower(CultureInfo.InvariantCulture));
			}
		}
	}

	public bool IncludeNewItemsInFilter
	{
		get
		{
			return GetXmlNodeBool("@includeNewItemsInFilter");
		}
		set
		{
			SetXmlNodeBool("@includeNewItemsInFilter", value);
		}
	}

	public eSubTotalFunctions SubTotalFunctions
	{
		get
		{
			eSubTotalFunctions eSubTotalFunctions2 = (eSubTotalFunctions)0;
			XmlNodeList xmlNodeList = base.TopNode.SelectNodes("d:items/d:item/@t", base.NameSpaceManager);
			if (xmlNodeList.Count == 0)
			{
				return eSubTotalFunctions.None;
			}
			foreach (XmlAttribute item in xmlNodeList)
			{
				try
				{
					eSubTotalFunctions2 |= (eSubTotalFunctions)Enum.Parse(typeof(eSubTotalFunctions), item.Value, ignoreCase: true);
				}
				catch (ArgumentException innerException)
				{
					throw new ArgumentException("Unable to parse value of " + item.Value + " to a valid pivot table subtotal function", innerException);
				}
			}
			return eSubTotalFunctions2;
		}
		set
		{
			if ((value & eSubTotalFunctions.None) == eSubTotalFunctions.None && value != eSubTotalFunctions.None)
			{
				throw new ArgumentException("Value None can not be combined with other values.");
			}
			if ((value & eSubTotalFunctions.Default) == eSubTotalFunctions.Default && value != eSubTotalFunctions.Default)
			{
				throw new ArgumentException("Value Default can not be combined with other values.");
			}
			XmlNodeList xmlNodeList = base.TopNode.SelectNodes("d:items/d:item/@t", base.NameSpaceManager);
			if (xmlNodeList.Count > 0)
			{
				foreach (XmlAttribute item in xmlNodeList)
				{
					DeleteNode("@" + item.Value + "Subtotal");
					item.OwnerElement.ParentNode.RemoveChild(item.OwnerElement);
				}
			}
			if (value == eSubTotalFunctions.None)
			{
				SetXmlNodeBool("@defaultSubtotal", value: false);
				base.TopNode.InnerXml = "";
				return;
			}
			string text = "";
			int num = 0;
			foreach (eSubTotalFunctions value2 in Enum.GetValues(typeof(eSubTotalFunctions)))
			{
				if ((value & value2) == value2)
				{
					string text2 = value2.ToString();
					string text3 = char.ToLowerInvariant(text2[0]) + text2.Substring(1);
					SetXmlNodeBool("@" + text3 + "Subtotal", value: true);
					text = text + "<item t=\"" + text3 + "\" />";
					num++;
				}
			}
			base.TopNode.InnerXml = $"<items count=\"{num}\">{text}</items>";
		}
	}

	public ePivotFieldAxis Axis
	{
		get
		{
			return GetXmlNodeString("@axis") switch
			{
				"axisRow" => ePivotFieldAxis.Row, 
				"axisCol" => ePivotFieldAxis.Column, 
				"axisPage" => ePivotFieldAxis.Page, 
				"axisValues" => ePivotFieldAxis.Values, 
				_ => ePivotFieldAxis.None, 
			};
		}
		internal set
		{
			switch (value)
			{
			case ePivotFieldAxis.Row:
				SetXmlNodeString("@axis", "axisRow");
				break;
			case ePivotFieldAxis.Column:
				SetXmlNodeString("@axis", "axisCol");
				break;
			case ePivotFieldAxis.Values:
				SetXmlNodeString("@axis", "axisValues");
				break;
			case ePivotFieldAxis.Page:
				SetXmlNodeString("@axis", "axisPage");
				break;
			default:
				DeleteNode("@axis");
				break;
			}
		}
	}

	public bool IsRowField
	{
		get
		{
			return base.TopNode.SelectSingleNode($"../../d:rowFields/d:field[@x={Index}]", base.NameSpaceManager) != null;
		}
		internal set
		{
			if (value)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode("../../d:rowFields", base.NameSpaceManager);
				if (xmlNode == null)
				{
					_table.CreateNode("d:rowFields");
				}
				xmlNode = base.TopNode.SelectSingleNode("../../d:rowFields", base.NameSpaceManager);
				AppendField(xmlNode, Index, "field", "x");
				if (BaseIndex == Index)
				{
					base.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
				}
				else
				{
					base.TopNode.InnerXml = "<items count=\"0\"></items>";
				}
			}
			else if (base.TopNode.SelectSingleNode($"../../d:rowFields/d:field[@x={Index}]", base.NameSpaceManager) is XmlElement xmlElement)
			{
				xmlElement.ParentNode.RemoveChild(xmlElement);
			}
		}
	}

	public bool IsColumnField
	{
		get
		{
			return base.TopNode.SelectSingleNode($"../../d:colFields/d:field[@x={Index}]", base.NameSpaceManager) != null;
		}
		internal set
		{
			if (value)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode("../../d:colFields", base.NameSpaceManager);
				if (xmlNode == null)
				{
					_table.CreateNode("d:colFields");
				}
				xmlNode = base.TopNode.SelectSingleNode("../../d:colFields", base.NameSpaceManager);
				AppendField(xmlNode, Index, "field", "x");
				if (BaseIndex == Index)
				{
					base.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
				}
				else
				{
					base.TopNode.InnerXml = "<items count=\"0\"></items>";
				}
			}
			else if (base.TopNode.SelectSingleNode($"../../d:colFields/d:field[@x={Index}]", base.NameSpaceManager) is XmlElement xmlElement)
			{
				xmlElement.ParentNode.RemoveChild(xmlElement);
			}
		}
	}

	public bool IsDataField => GetXmlNodeBool("@dataField", blankValue: false);

	public bool IsPageField
	{
		get
		{
			return Axis == ePivotFieldAxis.Page;
		}
		internal set
		{
			if (value)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode("../../d:pageFields", base.NameSpaceManager);
				if (xmlNode == null)
				{
					_table.CreateNode("d:pageFields");
					xmlNode = base.TopNode.SelectSingleNode("../../d:pageFields", base.NameSpaceManager);
				}
				base.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
				XmlElement topNode = AppendField(xmlNode, Index, "pageField", "fld");
				_pageFieldSettings = new ExcelPivotTablePageFieldSettings(base.NameSpaceManager, topNode, this, Index);
			}
			else
			{
				_pageFieldSettings = null;
				if (base.TopNode.SelectSingleNode($"../../d:pageFields/d:pageField[@fld={Index}]", base.NameSpaceManager) is XmlElement xmlElement)
				{
					xmlElement.ParentNode.RemoveChild(xmlElement);
				}
			}
		}
	}

	public ExcelPivotTablePageFieldSettings PageFieldSettings => _pageFieldSettings;

	internal eDateGroupBy DateGrouping { get; set; }

	public ExcelPivotTableFieldGroup Grouping => _grouping;

	public ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> Items
	{
		get
		{
			if (_items == null)
			{
				_items = new ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>(_table);
				foreach (XmlNode item in base.TopNode.SelectNodes("d:items//d:item", base.NameSpaceManager))
				{
					ExcelPivotTableFieldItem excelPivotTableFieldItem = new ExcelPivotTableFieldItem(base.NameSpaceManager, item, this);
					if (excelPivotTableFieldItem.T == "")
					{
						_items.AddInternal(excelPivotTableFieldItem);
					}
				}
			}
			return _items;
		}
	}

	internal ExcelPivotTableField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTable table, int index, int baseIndex)
		: base(ns, topNode)
	{
		Index = index;
		BaseIndex = baseIndex;
		_table = table;
	}

	internal XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
	{
		XmlElement refChild = null;
		foreach (XmlElement childNode in rowsNode.ChildNodes)
		{
			if (int.TryParse(childNode.GetAttribute(indexAttrText), out var result) && result == index)
			{
				return childNode;
			}
			refChild = childNode;
		}
		XmlElement xmlElement2 = rowsNode.OwnerDocument.CreateElement(fieldNodeText, "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement2.SetAttribute(indexAttrText, index.ToString());
		rowsNode.InsertAfter(xmlElement2, refChild);
		return xmlElement2;
	}

	internal void SetCacheFieldNode(XmlNode cacheField)
	{
		_cacheFieldHelper = new XmlHelperInstance(base.NameSpaceManager, cacheField);
		XmlNode xmlNode = cacheField.SelectSingleNode("d:fieldGroup", base.NameSpaceManager);
		if (xmlNode != null)
		{
			XmlNode xmlNode2 = xmlNode.SelectSingleNode("d:rangePr/@groupBy", base.NameSpaceManager);
			if (xmlNode2 == null)
			{
				_grouping = new ExcelPivotTableFieldNumericGroup(base.NameSpaceManager, cacheField);
				return;
			}
			DateGrouping = (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), xmlNode2.Value, ignoreCase: true);
			_grouping = new ExcelPivotTableFieldDateGroup(base.NameSpaceManager, xmlNode);
		}
	}

	internal ExcelPivotTableFieldDateGroup SetDateGroup(eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
	{
		ExcelPivotTableFieldDateGroup excelPivotTableFieldDateGroup = new ExcelPivotTableFieldDateGroup(base.NameSpaceManager, _cacheFieldHelper.TopNode);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsDate", value: true);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNonDate", value: false);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", value: false);
		excelPivotTableFieldDateGroup.TopNode.InnerXml += $"<fieldGroup base=\"{BaseIndex}\"><rangePr groupBy=\"{GroupBy.ToString().ToLower(CultureInfo.InvariantCulture)}\" /><groupItems /></fieldGroup>";
		if (StartDate.Year < 1900)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", "1900-01-01T00:00:00");
		}
		else
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoStart", "0");
		}
		if (EndDate == DateTime.MaxValue)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", "9999-12-31T00:00:00");
		}
		else
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
		}
		int items = AddDateGroupItems(excelPivotTableFieldDateGroup, GroupBy, StartDate, EndDate, interval);
		AddFieldItems(items);
		_grouping = excelPivotTableFieldDateGroup;
		return excelPivotTableFieldDateGroup;
	}

	internal ExcelPivotTableFieldNumericGroup SetNumericGroup(double start, double end, double interval)
	{
		ExcelPivotTableFieldNumericGroup excelPivotTableFieldNumericGroup = new ExcelPivotTableFieldNumericGroup(base.NameSpaceManager, _cacheFieldHelper.TopNode);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNumber", value: true);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsInteger", value: true);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", value: false);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsString", value: false);
		excelPivotTableFieldNumericGroup.TopNode.InnerXml += $"<fieldGroup base=\"{BaseIndex}\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{start.ToString(CultureInfo.InvariantCulture)}\" endNum=\"{end.ToString(CultureInfo.InvariantCulture)}\" groupInterval=\"{interval.ToString(CultureInfo.InvariantCulture)}\"/><groupItems /></fieldGroup>";
		int items = AddNumericGroupItems(excelPivotTableFieldNumericGroup, start, end, interval);
		AddFieldItems(items);
		_grouping = excelPivotTableFieldNumericGroup;
		return excelPivotTableFieldNumericGroup;
	}

	private int AddNumericGroupItems(ExcelPivotTableFieldNumericGroup group, double start, double end, double interval)
	{
		if (interval < 0.0)
		{
			throw new Exception("The interval must be a positiv");
		}
		if (start > end)
		{
			throw new Exception("Then End number must be larger than the Start number");
		}
		XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
		int num = 2;
		double num2 = start;
		double num3 = start + interval;
		AddGroupItem(groupItems, "<" + start.ToString(CultureInfo.InvariantCulture));
		while (num2 < end)
		{
			AddGroupItem(groupItems, $"{num2.ToString(CultureInfo.InvariantCulture)}-{num3.ToString(CultureInfo.InvariantCulture)}");
			num2 = num3;
			num3 += interval;
			num++;
		}
		AddGroupItem(groupItems, ">" + num3.ToString(CultureInfo.InvariantCulture));
		return num;
	}

	private void AddFieldItems(int items)
	{
		XmlElement xmlElement = null;
		XmlElement xmlElement2 = base.TopNode.SelectSingleNode("d:items", base.NameSpaceManager) as XmlElement;
		for (int i = 0; i < items; i++)
		{
			XmlElement xmlElement3 = xmlElement2.OwnerDocument.CreateElement("item", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement3.SetAttribute("x", i.ToString());
			if (xmlElement == null)
			{
				xmlElement2.PrependChild(xmlElement3);
			}
			else
			{
				xmlElement2.InsertAfter(xmlElement3, xmlElement);
			}
			xmlElement = xmlElement3;
		}
		xmlElement2.SetAttribute("count", (items + 1).ToString());
	}

	private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
	{
		XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
		int num = 2;
		AddGroupItem(groupItems, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));
		switch (GroupBy)
		{
		case eDateGroupBy.Minutes:
		case eDateGroupBy.Seconds:
			AddTimeSerie(60, groupItems);
			num += 60;
			break;
		case eDateGroupBy.Hours:
			AddTimeSerie(24, groupItems);
			num += 24;
			break;
		case eDateGroupBy.Days:
		{
			if (interval == 1)
			{
				DateTime dateTime = new DateTime(2008, 1, 1);
				while (dateTime.Year == 2008)
				{
					AddGroupItem(groupItems, dateTime.ToString("dd-MMM"));
					dateTime = dateTime.AddDays(1.0);
				}
				num += 366;
				break;
			}
			DateTime dateTime2 = StartDate;
			num = 0;
			while (dateTime2 < EndDate)
			{
				AddGroupItem(groupItems, dateTime2.ToString("dd-MMM"));
				dateTime2 = dateTime2.AddDays(interval);
				num++;
			}
			break;
		}
		case eDateGroupBy.Months:
			AddGroupItem(groupItems, "jan");
			AddGroupItem(groupItems, "feb");
			AddGroupItem(groupItems, "mar");
			AddGroupItem(groupItems, "apr");
			AddGroupItem(groupItems, "may");
			AddGroupItem(groupItems, "jun");
			AddGroupItem(groupItems, "jul");
			AddGroupItem(groupItems, "aug");
			AddGroupItem(groupItems, "sep");
			AddGroupItem(groupItems, "oct");
			AddGroupItem(groupItems, "nov");
			AddGroupItem(groupItems, "dec");
			num += 12;
			break;
		case eDateGroupBy.Quarters:
			AddGroupItem(groupItems, "Qtr1");
			AddGroupItem(groupItems, "Qtr2");
			AddGroupItem(groupItems, "Qtr3");
			AddGroupItem(groupItems, "Qtr4");
			num += 4;
			break;
		case eDateGroupBy.Years:
			if (StartDate.Year >= 1900 && EndDate != DateTime.MaxValue)
			{
				for (int i = StartDate.Year; i <= EndDate.Year; i++)
				{
					AddGroupItem(groupItems, i.ToString());
				}
				num += EndDate.Year - StartDate.Year + 1;
			}
			break;
		default:
			throw new Exception("unsupported grouping");
		}
		AddGroupItem(groupItems, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));
		return num;
	}

	private void AddTimeSerie(int count, XmlElement groupItems)
	{
		for (int i = 0; i < count; i++)
		{
			AddGroupItem(groupItems, $"{i:00}");
		}
	}

	private void AddGroupItem(XmlElement groupItems, string value)
	{
		XmlElement xmlElement = groupItems.OwnerDocument.CreateElement("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("v", value);
		groupItems.AppendChild(xmlElement);
	}

	public void AddNumericGrouping(double Start, double End, double Interval)
	{
		ValidateGrouping();
		SetNumericGroup(Start, End, Interval);
	}

	public void AddDateGrouping(eDateGroupBy groupBy)
	{
		AddDateGrouping(groupBy, DateTime.MinValue, DateTime.MaxValue, 1);
	}

	public void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate)
	{
		AddDateGrouping(groupBy, startDate, endDate, 1);
	}

	public void AddDateGrouping(int days, DateTime startDate, DateTime endDate)
	{
		AddDateGrouping(eDateGroupBy.Days, startDate, endDate, days);
	}

	private void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int groupInterval)
	{
		if (groupInterval < 1 || groupInterval >= 32767)
		{
			throw new ArgumentOutOfRangeException("Group interval is out of range");
		}
		if (groupInterval > 1 && groupBy != eDateGroupBy.Days)
		{
			throw new ArgumentException("Group interval is can only be used when groupBy is Days");
		}
		ValidateGrouping();
		bool firstField = true;
		List<ExcelPivotTableField> list = new List<ExcelPivotTableField>();
		if ((groupBy & eDateGroupBy.Seconds) == eDateGroupBy.Seconds)
		{
			list.Add(AddField(eDateGroupBy.Seconds, startDate, endDate, ref firstField));
		}
		if ((groupBy & eDateGroupBy.Minutes) == eDateGroupBy.Minutes)
		{
			list.Add(AddField(eDateGroupBy.Minutes, startDate, endDate, ref firstField));
		}
		if ((groupBy & eDateGroupBy.Hours) == eDateGroupBy.Hours)
		{
			list.Add(AddField(eDateGroupBy.Hours, startDate, endDate, ref firstField));
		}
		if ((groupBy & eDateGroupBy.Days) == eDateGroupBy.Days)
		{
			list.Add(AddField(eDateGroupBy.Days, startDate, endDate, ref firstField, groupInterval));
		}
		if ((groupBy & eDateGroupBy.Months) == eDateGroupBy.Months)
		{
			list.Add(AddField(eDateGroupBy.Months, startDate, endDate, ref firstField));
		}
		if ((groupBy & eDateGroupBy.Quarters) == eDateGroupBy.Quarters)
		{
			list.Add(AddField(eDateGroupBy.Quarters, startDate, endDate, ref firstField));
		}
		if ((groupBy & eDateGroupBy.Years) == eDateGroupBy.Years)
		{
			list.Add(AddField(eDateGroupBy.Years, startDate, endDate, ref firstField));
		}
		if (list.Count > 1)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/@par", (_table.Fields.Count - 1).ToString());
		}
		if (groupInterval != 1)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", groupInterval.ToString());
		}
		else
		{
			_cacheFieldHelper.DeleteNode("d:fieldGroup/d:rangePr/@groupInterval");
		}
		_items = null;
	}

	private void ValidateGrouping()
	{
		if (!IsColumnField && !IsRowField)
		{
			throw new Exception("Field must be a row or column field");
		}
		foreach (ExcelPivotTableField field in _table.Fields)
		{
			if (field.Grouping != null)
			{
				throw new Exception("Grouping already exists");
			}
		}
	}

	private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField)
	{
		return AddField(groupBy, startDate, endDate, ref firstField, 1);
	}

	private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField, int interval)
	{
		if (!firstField)
		{
			XmlNode xmlNode = _table.PivotTableXml.SelectSingleNode("//d:pivotFields", _table.NameSpaceManager);
			XmlElement xmlElement = _table.PivotTableXml.CreateElement("pivotField", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.SetAttribute("compact", "0");
			xmlElement.SetAttribute("outline", "0");
			xmlElement.SetAttribute("showAll", "0");
			xmlElement.SetAttribute("defaultSubtotal", "0");
			xmlNode.AppendChild(xmlElement);
			ExcelPivotTableField excelPivotTableField = new ExcelPivotTableField(_table.NameSpaceManager, xmlElement, _table, _table.Fields.Count, Index);
			excelPivotTableField.DateGrouping = groupBy;
			XmlNode xmlNode2 = ((!IsRowField) ? base.TopNode.SelectSingleNode("../../d:colFields", base.NameSpaceManager) : base.TopNode.SelectSingleNode("../../d:rowFields", base.NameSpaceManager));
			int num = 0;
			foreach (XmlElement childNode in xmlNode2.ChildNodes)
			{
				if (int.TryParse(childNode.GetAttribute("x"), out var result) && _table.Fields[result].BaseIndex == BaseIndex)
				{
					XmlElement xmlElement3 = xmlNode2.OwnerDocument.CreateElement("field", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					xmlElement3.SetAttribute("x", excelPivotTableField.Index.ToString());
					xmlNode2.InsertBefore(xmlElement3, childNode);
					break;
				}
				num++;
			}
			if (IsRowField)
			{
				_table.RowFields.Insert(excelPivotTableField, num);
			}
			else
			{
				_table.ColumnFields.Insert(excelPivotTableField, num);
			}
			_table.Fields.AddInternal(excelPivotTableField);
			AddCacheField(excelPivotTableField, startDate, endDate, interval);
			return excelPivotTableField;
		}
		firstField = false;
		DateGrouping = groupBy;
		Compact = false;
		SetDateGroup(groupBy, startDate, endDate, interval);
		return this;
	}

	private void AddCacheField(ExcelPivotTableField field, DateTime startDate, DateTime endDate, int interval)
	{
		XmlNode xmlNode = _table.CacheDefinition.CacheDefinitionXml.SelectSingleNode("//d:cacheFields", _table.NameSpaceManager);
		XmlElement xmlElement = _table.CacheDefinition.CacheDefinitionXml.CreateElement("cacheField", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("name", field.DateGrouping.ToString());
		xmlElement.SetAttribute("databaseField", "0");
		xmlNode.AppendChild(xmlElement);
		field.SetCacheFieldNode(xmlElement);
		field.SetDateGroup(field.DateGrouping, startDate, endDate, interval);
	}
}
