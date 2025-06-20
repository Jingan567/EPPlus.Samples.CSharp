using System;
using System.Collections.Generic;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table;

public class ExcelTable : XmlHelper, IEqualityComparer<ExcelTable>
{
	private const string ID_PATH = "@id";

	private const string NAME_PATH = "@name";

	private const string DISPLAY_NAME_PATH = "@displayName";

	private ExcelAddressBase _address;

	internal ExcelTableColumnCollection _cols;

	private TableStyles _tableStyle = TableStyles.Medium6;

	private const string HEADERROWCOUNT_PATH = "@headerRowCount";

	private const string AUTOFILTER_PATH = "d:autoFilter/@ref";

	private const string TOTALSROWCOUNT_PATH = "@totalsRowCount";

	private const string TOTALSROWSHOWN_PATH = "@totalsRowShown";

	private const string STYLENAME_PATH = "d:tableStyleInfo/@name";

	private const string SHOWFIRSTCOLUMN_PATH = "d:tableStyleInfo/@showFirstColumn";

	private const string SHOWLASTCOLUMN_PATH = "d:tableStyleInfo/@showLastColumn";

	private const string SHOWROWSTRIPES_PATH = "d:tableStyleInfo/@showRowStripes";

	private const string SHOWCOLUMNSTRIPES_PATH = "d:tableStyleInfo/@showColumnStripes";

	private const string TOTALSROWCELLSTYLE_PATH = "@totalsRowCellStyle";

	private const string DATACELLSTYLE_PATH = "@dataCellStyle";

	private const string HEADERROWCELLSTYLE_PATH = "@headerRowCellStyle";

	internal ZipPackagePart Part { get; set; }

	public XmlDocument TableXml { get; set; }

	public Uri TableUri { get; internal set; }

	internal string RelationshipID { get; set; }

	internal int Id
	{
		get
		{
			return GetXmlNodeInt("@id");
		}
		set
		{
			SetXmlNodeString("@id", value.ToString());
		}
	}

	public string Name
	{
		get
		{
			return GetXmlNodeString("@name");
		}
		set
		{
			if (!Name.Equals(value, StringComparison.CurrentCultureIgnoreCase) && WorkSheet.Workbook.ExistsTableName(value))
			{
				throw new ArgumentException("Tablename is not unique");
			}
			string name = Name;
			if (WorkSheet.Tables._tableNames.ContainsKey(name))
			{
				int value2 = WorkSheet.Tables._tableNames[name];
				WorkSheet.Tables._tableNames.Remove(name);
				WorkSheet.Tables._tableNames.Add(value, value2);
			}
			SetXmlNodeString("@name", value);
			SetXmlNodeString("@displayName", cleanDisplayName(value));
		}
	}

	public ExcelWorksheet WorkSheet { get; set; }

	public ExcelAddressBase Address
	{
		get
		{
			return _address;
		}
		internal set
		{
			_address = value;
			SetXmlNodeString("@ref", value.Address);
			WriteAutoFilter(ShowTotal);
		}
	}

	public ExcelTableColumnCollection Columns
	{
		get
		{
			if (_cols == null)
			{
				_cols = new ExcelTableColumnCollection(this);
			}
			return _cols;
		}
	}

	public TableStyles TableStyle
	{
		get
		{
			return _tableStyle;
		}
		set
		{
			_tableStyle = value;
			if (value != TableStyles.Custom)
			{
				SetXmlNodeString("d:tableStyleInfo/@name", "TableStyle" + value);
			}
		}
	}

	public bool ShowHeader
	{
		get
		{
			return GetXmlNodeInt("@headerRowCount") != 0;
		}
		set
		{
			if ((Address._toRow - Address._fromRow < 0 && value) || (Address._toRow - Address._fromRow == 1 && value && ShowTotal))
			{
				throw new Exception("Cant set ShowHeader-property. Table has too few rows");
			}
			if (value)
			{
				DeleteNode("@headerRowCount");
				WriteAutoFilter(ShowTotal);
				for (int i = 0; i < Columns.Count; i++)
				{
					string value2 = WorkSheet.GetValue<string>(Address._fromRow, Address._fromCol + i);
					if (string.IsNullOrEmpty(value2))
					{
						WorkSheet.SetValue(Address._fromRow, Address._fromCol + i, _cols[i].Name);
					}
					else if (value2 != _cols[i].Name)
					{
						_cols[i].Name = value2;
					}
				}
			}
			else
			{
				SetXmlNodeString("@headerRowCount", "0");
				DeleteAllNode("d:autoFilter/@ref");
			}
		}
	}

	internal ExcelAddressBase AutoFilterAddress
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:autoFilter/@ref");
			if (xmlNodeString == "")
			{
				return null;
			}
			return new ExcelAddressBase(xmlNodeString);
		}
	}

	public bool ShowFilter
	{
		get
		{
			if (ShowHeader)
			{
				return AutoFilterAddress != null;
			}
			return false;
		}
		set
		{
			if (ShowHeader)
			{
				if (value)
				{
					WriteAutoFilter(ShowTotal);
				}
				else
				{
					DeleteAllNode("d:autoFilter/@ref");
				}
			}
			else if (value)
			{
				throw new InvalidOperationException("Filter can only be applied when ShowHeader is set to true");
			}
		}
	}

	public bool ShowTotal
	{
		get
		{
			return GetXmlNodeInt("@totalsRowCount") == 1;
		}
		set
		{
			if (value != ShowTotal)
			{
				if (value)
				{
					Address = new ExcelAddress(WorkSheet.Name, ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column, Address.End.Row + 1, Address.End.Column));
				}
				else
				{
					Address = new ExcelAddress(WorkSheet.Name, ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column, Address.End.Row - 1, Address.End.Column));
				}
				SetXmlNodeString("@ref", Address.Address);
				if (value)
				{
					SetXmlNodeString("@totalsRowCount", "1");
				}
				else
				{
					DeleteNode("@totalsRowCount");
				}
				WriteAutoFilter(value);
			}
		}
	}

	public string StyleName
	{
		get
		{
			return GetXmlNodeString("d:tableStyleInfo/@name");
		}
		set
		{
			if (value.StartsWith("TableStyle"))
			{
				try
				{
					_tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), ignoreCase: true);
				}
				catch
				{
					_tableStyle = TableStyles.Custom;
				}
			}
			else if (value == "None")
			{
				_tableStyle = TableStyles.None;
				value = "";
			}
			else
			{
				_tableStyle = TableStyles.Custom;
			}
			SetXmlNodeString("d:tableStyleInfo/@name", value, removeIfBlank: true);
		}
	}

	public bool ShowFirstColumn
	{
		get
		{
			return GetXmlNodeBool("d:tableStyleInfo/@showFirstColumn");
		}
		set
		{
			SetXmlNodeBool("d:tableStyleInfo/@showFirstColumn", value, removeIf: false);
		}
	}

	public bool ShowLastColumn
	{
		get
		{
			return GetXmlNodeBool("d:tableStyleInfo/@showLastColumn");
		}
		set
		{
			SetXmlNodeBool("d:tableStyleInfo/@showLastColumn", value, removeIf: false);
		}
	}

	public bool ShowRowStripes
	{
		get
		{
			return GetXmlNodeBool("d:tableStyleInfo/@showRowStripes");
		}
		set
		{
			SetXmlNodeBool("d:tableStyleInfo/@showRowStripes", value, removeIf: false);
		}
	}

	public bool ShowColumnStripes
	{
		get
		{
			return GetXmlNodeBool("d:tableStyleInfo/@showColumnStripes");
		}
		set
		{
			SetXmlNodeBool("d:tableStyleInfo/@showColumnStripes", value, removeIf: false);
		}
	}

	public string TotalsRowCellStyle
	{
		get
		{
			return GetXmlNodeString("@totalsRowCellStyle");
		}
		set
		{
			if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
			{
				throw new Exception($"Named style {value} does not exist.");
			}
			SetXmlNodeString(base.TopNode, "@totalsRowCellStyle", value, removeIfBlank: true);
			if (ShowTotal)
			{
				WorkSheet.Cells[Address._toRow, Address._fromCol, Address._toRow, Address._toCol].StyleName = value;
			}
		}
	}

	public string DataCellStyleName
	{
		get
		{
			return GetXmlNodeString("@dataCellStyle");
		}
		set
		{
			if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
			{
				throw new Exception($"Named style {value} does not exist.");
			}
			SetXmlNodeString(base.TopNode, "@dataCellStyle", value, removeIfBlank: true);
			int num = Address._fromRow + (ShowHeader ? 1 : 0);
			int num2 = Address._toRow - (ShowTotal ? 1 : 0);
			if (num < num2)
			{
				WorkSheet.Cells[num, Address._fromCol, num2, Address._toCol].StyleName = value;
			}
		}
	}

	public string HeaderRowCellStyle
	{
		get
		{
			return GetXmlNodeString("@headerRowCellStyle");
		}
		set
		{
			if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
			{
				throw new Exception($"Named style {value} does not exist.");
			}
			SetXmlNodeString(base.TopNode, "@headerRowCellStyle", value, removeIfBlank: true);
			if (ShowHeader)
			{
				WorkSheet.Cells[Address._fromRow, Address._fromCol, Address._fromRow, Address._toCol].StyleName = value;
			}
		}
	}

	internal ExcelTable(ZipPackageRelationship rel, ExcelWorksheet sheet)
		: base(sheet.NameSpaceManager)
	{
		WorkSheet = sheet;
		TableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
		RelationshipID = rel.Id;
		ZipPackage package = sheet._package.Package;
		Part = package.GetPart(TableUri);
		TableXml = new XmlDocument();
		XmlHelper.LoadXmlSafe(TableXml, Part.GetStream());
		init();
		Address = new ExcelAddressBase(GetXmlNodeString("@ref"));
	}

	internal ExcelTable(ExcelWorksheet sheet, ExcelAddressBase address, string name, int tblId)
		: base(sheet.NameSpaceManager)
	{
		WorkSheet = sheet;
		Address = address;
		TableXml = new XmlDocument();
		XmlHelper.LoadXmlSafe(TableXml, GetStartXml(name, tblId), Encoding.UTF8);
		base.TopNode = TableXml.DocumentElement;
		init();
		if (address._fromRow == address._toRow)
		{
			ShowHeader = false;
		}
	}

	private void init()
	{
		base.TopNode = TableXml.DocumentElement;
		base.SchemaNodeOrder = new string[3] { "autoFilter", "tableColumns", "tableStyleInfo" };
	}

	private string GetStartXml(string name, int tblId)
	{
		name = ConvertUtil.ExcelEscapeString(name);
		string text = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>";
		text += $"<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{tblId}\" name=\"{name}\" displayName=\"{cleanDisplayName(name)}\" ref=\"{Address.Address}\" headerRowCount=\"1\">";
		text += $"<autoFilter ref=\"{Address.Address}\" />";
		int num = Address._toCol - Address._fromCol + 1;
		text += $"<tableColumns count=\"{num}\">";
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		for (int i = 1; i <= num; i++)
		{
			ExcelRange excelRange = WorkSheet.Cells[Address._fromRow, Address._fromCol + i - 1];
			string text2;
			if (excelRange.Value == null || dictionary.ContainsKey(excelRange.Value.ToString()))
			{
				int num2 = i;
				do
				{
					text2 = $"Column{num2++}";
				}
				while (dictionary.ContainsKey(text2));
			}
			else
			{
				text2 = SecurityElement.Escape(excelRange.Value.ToString());
			}
			dictionary.Add(text2, text2);
			text += $"<tableColumn id=\"{i}\" name=\"{text2}\" />";
		}
		text += "</tableColumns>";
		text += "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" /> ";
		return text + "</table>";
	}

	private string cleanDisplayName(string name)
	{
		return Regex.Replace(name, "[^\\w\\.-_]", "_");
	}

	private void WriteAutoFilter(bool showTotal)
	{
		if (ShowHeader)
		{
			string value = ((!showTotal) ? Address.Address : ExcelCellBase.GetAddress(Address._fromRow, Address._fromCol, Address._toRow - 1, Address._toCol));
			SetXmlNodeString("d:autoFilter/@ref", value);
		}
	}

	public bool Equals(ExcelTable x, ExcelTable y)
	{
		if (x.WorkSheet == y.WorkSheet && x.Id == y.Id)
		{
			return x.TableXml.OuterXml == y.TableXml.OuterXml;
		}
		return false;
	}

	public int GetHashCode(ExcelTable obj)
	{
		return obj.TableXml.OuterXml.GetHashCode();
	}
}
