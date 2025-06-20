using System;
using System.Globalization;

namespace OfficeOpenXml.Style.Dxf;

public abstract class DxfStyleBase<T>
{
	protected ExcelStyles _styles;

	protected internal abstract string Id { get; }

	protected internal abstract bool HasValue { get; }

	protected internal bool AllowChange { get; set; }

	internal DxfStyleBase(ExcelStyles styles)
	{
		_styles = styles;
		AllowChange = false;
	}

	protected internal abstract void CreateNodes(XmlHelper helper, string path);

	protected internal abstract T Clone();

	protected void SetValueColor(XmlHelper helper, string path, ExcelDxfColor color)
	{
		if (color != null && color.HasValue)
		{
			if (color.Color.HasValue)
			{
				SetValue(helper, path + "/@rgb", color.Color.Value.ToArgb().ToString("x"));
			}
			else if (color.Auto.HasValue)
			{
				SetValueBool(helper, path + "/@auto", color.Auto);
			}
			else if (color.Theme.HasValue)
			{
				SetValue(helper, path + "/@theme", color.Theme);
			}
			else if (color.Index.HasValue)
			{
				SetValue(helper, path + "/@indexed", color.Index);
			}
			if (color.Tint.HasValue)
			{
				SetValue(helper, path + "/@tint", color.Tint);
			}
		}
	}

	protected void SetValueEnum(XmlHelper helper, string path, Enum v)
	{
		if (v == null)
		{
			helper.DeleteNode(path);
			return;
		}
		string text = v.ToString();
		text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1);
		helper.SetXmlNodeString(path, text);
	}

	protected void SetValue(XmlHelper helper, string path, object v)
	{
		if (v == null)
		{
			helper.DeleteNode(path);
		}
		else
		{
			helper.SetXmlNodeString(path, v.ToString());
		}
	}

	protected void SetValueBool(XmlHelper helper, string path, bool? v)
	{
		if (!v.HasValue)
		{
			helper.DeleteNode(path);
		}
		else
		{
			helper.SetXmlNodeBool(path, v.Value);
		}
	}

	protected internal string GetAsString(object v)
	{
		return (v ?? "").ToString();
	}
}
