using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableDataFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableDataField>
{
	internal ExcelPivotTableDataFieldCollection(ExcelPivotTable table)
		: base(table)
	{
	}

	public ExcelPivotTableDataField Add(ExcelPivotTableField field)
	{
		XmlNode xmlNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
		if (xmlNode == null)
		{
			_table.CreateNode("d:dataFields");
			xmlNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
		}
		XmlElement xmlElement = _table.PivotTableXml.CreateElement("dataField", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("fld", field.Index.ToString());
		xmlNode.AppendChild(xmlElement);
		field.SetXmlNodeBool("@dataField", value: true, removeIf: false);
		ExcelPivotTableDataField excelPivotTableDataField = new ExcelPivotTableDataField(field.NameSpaceManager, xmlElement, field);
		ValidateDupName(excelPivotTableDataField);
		_list.Add(excelPivotTableDataField);
		return excelPivotTableDataField;
	}

	private void ValidateDupName(ExcelPivotTableDataField dataField)
	{
		if (ExistsDfName(dataField.Field.Name, null))
		{
			int num = 2;
			string name;
			do
			{
				name = dataField.Field.Name + "_" + num++;
			}
			while (ExistsDfName(name, null));
			dataField.Name = name;
		}
	}

	internal bool ExistsDfName(string name, ExcelPivotTableDataField datafield)
	{
		foreach (ExcelPivotTableDataField item in _list)
		{
			if (((!string.IsNullOrEmpty(item.Name) && item.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) || (string.IsNullOrEmpty(item.Name) && item.Field.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) && datafield != item)
			{
				return true;
			}
		}
		return false;
	}

	public void Remove(ExcelPivotTableDataField dataField)
	{
		if (dataField.Field.TopNode.SelectSingleNode($"../../d:dataFields/d:dataField[@fld={dataField.Index}]", dataField.NameSpaceManager) is XmlElement xmlElement)
		{
			xmlElement.ParentNode.RemoveChild(xmlElement);
		}
		_list.Remove(dataField);
	}
}
