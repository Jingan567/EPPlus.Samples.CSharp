using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingHelper
{
	public static string CheckAndFixRangeAddress(string address)
	{
		if (address.Contains(','))
		{
			throw new FormatException("Multiple addresses may not be commaseparated, use space instead");
		}
		address = ConvertUtil._invariantTextInfo.ToUpper(address);
		if (Regex.IsMatch(address, "[A-Z]+:[A-Z]+"))
		{
			address = AddressUtility.ParseEntireColumnSelections(address);
		}
		return address;
	}

	public static Color ConvertFromColorCode(string colorCode)
	{
		try
		{
			return Color.FromArgb(int.Parse(colorCode.Replace("#", ""), NumberStyles.HexNumber));
		}
		catch
		{
			return Color.White;
		}
	}

	public static string GetAttributeString(XmlNode node, string attribute)
	{
		try
		{
			string value = node.Attributes[attribute].Value;
			return (value == null) ? string.Empty : value;
		}
		catch
		{
			return string.Empty;
		}
	}

	public static int GetAttributeInt(XmlNode node, string attribute)
	{
		try
		{
			return int.Parse(node.Attributes[attribute].Value, NumberStyles.Integer, CultureInfo.InvariantCulture);
		}
		catch
		{
			return int.MinValue;
		}
	}

	public static int? GetAttributeIntNullable(XmlNode node, string attribute)
	{
		try
		{
			if (node.Attributes[attribute] == null)
			{
				return null;
			}
			return int.Parse(node.Attributes[attribute].Value, NumberStyles.Integer, CultureInfo.InvariantCulture);
		}
		catch
		{
			return null;
		}
	}

	public static bool GetAttributeBool(XmlNode node, string attribute)
	{
		try
		{
			string value = node.Attributes[attribute].Value;
			return value == "1" || value == "-1" || value.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
		}
		catch
		{
			return false;
		}
	}

	public static bool? GetAttributeBoolNullable(XmlNode node, string attribute)
	{
		try
		{
			if (node.Attributes[attribute] == null)
			{
				return null;
			}
			string value = node.Attributes[attribute].Value;
			return value == "1" || value == "-1" || value.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
		}
		catch
		{
			return null;
		}
	}

	public static double GetAttributeDouble(XmlNode node, string attribute)
	{
		try
		{
			return double.Parse(node.Attributes[attribute].Value, NumberStyles.Number, CultureInfo.InvariantCulture);
		}
		catch
		{
			return double.NaN;
		}
	}

	public static decimal GetAttributeDecimal(XmlNode node, string attribute)
	{
		try
		{
			return decimal.Parse(node.Attributes[attribute].Value, NumberStyles.Any, CultureInfo.InvariantCulture);
		}
		catch
		{
			return decimal.MinValue;
		}
	}

	public static string EncodeXML(this string s)
	{
		return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
			.Replace("\"", "&quot;")
			.Replace("'", "&apos;");
	}

	public static string DecodeXML(this string s)
	{
		return s.Replace("'", "&apos;").Replace("\"", "&quot;").Replace(">", "&gt;")
			.Replace("<", "&lt;")
			.Replace("&", "&amp;");
	}
}
