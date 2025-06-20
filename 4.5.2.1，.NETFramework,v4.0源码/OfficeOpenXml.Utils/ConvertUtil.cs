using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml.Utils;

internal static class ConvertUtil
{
	internal static TextInfo _invariantTextInfo = CultureInfo.InvariantCulture.TextInfo;

	internal static CompareInfo _invariantCompareInfo = CompareInfo.GetCompareInfo(CultureInfo.InvariantCulture.Name);

	internal static bool IsNumeric(object candidate)
	{
		if (candidate == null)
		{
			return false;
		}
		if (!TypeCompat.IsPrimitive(candidate) && !(candidate is double) && !(candidate is decimal) && !(candidate is DateTime) && !(candidate is TimeSpan))
		{
			return candidate is long;
		}
		return true;
	}

	internal static bool TryParseNumericString(object candidate, out double result)
	{
		if (candidate != null)
		{
			NumberStyles style = NumberStyles.Float | NumberStyles.AllowThousands;
			return double.TryParse(candidate.ToString(), style, CultureInfo.CurrentCulture, out result);
		}
		result = 0.0;
		return false;
	}

	internal static bool TryParseBooleanString(object candidate, out bool result)
	{
		if (candidate != null)
		{
			return bool.TryParse(candidate.ToString(), out result);
		}
		result = false;
		return false;
	}

	internal static bool TryParseDateString(object candidate, out DateTime result)
	{
		if (candidate != null)
		{
			DateTimeStyles styles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal;
			return DateTime.TryParse(candidate.ToString(), CultureInfo.CurrentCulture, styles, out result);
		}
		result = DateTime.MinValue;
		return false;
	}

	internal static double GetValueDouble(object v, bool ignoreBool = false, bool retNaN = false)
	{
		try
		{
			if (ignoreBool && v is bool)
			{
				return 0.0;
			}
			if (IsNumeric(v))
			{
				if (v is DateTime dateTime)
				{
					return dateTime.ToOADate();
				}
				if (v is TimeSpan)
				{
					return DateTime.FromOADate(0.0).Add((TimeSpan)v).ToOADate();
				}
				return Convert.ToDouble(v, CultureInfo.InvariantCulture);
			}
			return retNaN ? double.NaN : 0.0;
		}
		catch
		{
			return retNaN ? double.NaN : 0.0;
		}
	}

	internal static string ExcelEscapeString(string s)
	{
		return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
	}

	internal static void ExcelEncodeString(StreamWriter sw, string t)
	{
		if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
		{
			Match match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
			int num = 0;
			while (match.Success)
			{
				t = t.Insert(match.Index + num, "_x005F");
				num += 6;
				match = match.NextMatch();
			}
		}
		for (int i = 0; i < t.Length; i++)
		{
			if (t[i] <= '\u001f' && t[i] != '\t' && t[i] != '\n' && t[i] != '\r')
			{
				sw.Write("_x00{0}_", ((t[i] < '\u000f') ? "0" : "") + ((int)t[i]).ToString("X"));
			}
			else
			{
				sw.Write(t[i]);
			}
		}
	}

	internal static void ExcelEncodeString(StringBuilder sb, string t, bool encodeTabCRLF = false)
	{
		if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
		{
			Match match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
			int num = 0;
			while (match.Success)
			{
				t = t.Insert(match.Index + num, "_x005F");
				num += 6;
				match = match.NextMatch();
			}
		}
		for (int i = 0; i < t.Length; i++)
		{
			if (t[i] <= '\u001f' && ((t[i] != '\t' && t[i] != '\n' && t[i] != '\r' && !encodeTabCRLF) || encodeTabCRLF))
			{
				sb.AppendFormat("_x00{0}_", ((t[i] < '\u000f') ? "0" : "") + ((int)t[i]).ToString("X"));
			}
			else
			{
				sb.Append(t[i]);
			}
		}
	}

	internal static string ExcelEncodeString(string t)
	{
		StringBuilder stringBuilder = new StringBuilder();
		t = t.Replace("\r\n", "\n");
		ExcelEncodeString(stringBuilder, t, encodeTabCRLF: true);
		return stringBuilder.ToString();
	}

	internal static string ExcelDecodeString(string t)
	{
		Match match = Regex.Match(t, "(_x005F|_x[0-9A-F]{4,4}_)");
		if (!match.Success)
		{
			return t;
		}
		bool flag = false;
		StringBuilder stringBuilder = new StringBuilder();
		int num = 0;
		while (match.Success)
		{
			if (num < match.Index)
			{
				stringBuilder.Append(t.Substring(num, match.Index - num));
			}
			if (!flag && match.Value == "_x005F")
			{
				flag = true;
			}
			else if (flag)
			{
				stringBuilder.Append(match.Value);
				flag = false;
			}
			else
			{
				stringBuilder.Append((char)int.Parse(match.Value.Substring(2, 4), NumberStyles.AllowHexSpecifier));
			}
			num = match.Index + match.Length;
			match = match.NextMatch();
		}
		stringBuilder.Append(t.Substring(num, t.Length - num));
		return stringBuilder.ToString();
	}

	public static T GetTypedCellValue<T>(object value)
	{
		if (value == null)
		{
			return default(T);
		}
		Type type = value.GetType();
		Type typeFromHandle = typeof(T);
		Type type2 = ((TypeCompat.IsGenericType(typeFromHandle) && typeFromHandle.GetGenericTypeDefinition() == typeof(Nullable<>)) ? Nullable.GetUnderlyingType(typeFromHandle) : null);
		if (type == typeFromHandle || type == type2)
		{
			return (T)value;
		}
		if (type2 != null && type == typeof(string) && ((string)value).Trim() == string.Empty)
		{
			return default(T);
		}
		typeFromHandle = type2 ?? typeFromHandle;
		if (typeFromHandle == typeof(DateTime))
		{
			if (value is double)
			{
				return (T)(object)DateTime.FromOADate((double)value);
			}
			if (type == typeof(TimeSpan))
			{
				return (T)(object)new DateTime(((TimeSpan)value).Ticks);
			}
			if (type == typeof(string))
			{
				return (T)(object)DateTime.Parse(value.ToString());
			}
		}
		else if (typeFromHandle == typeof(TimeSpan))
		{
			if (value is double)
			{
				return (T)(object)new TimeSpan(DateTime.FromOADate((double)value).Ticks);
			}
			if (type == typeof(DateTime))
			{
				return (T)(object)new TimeSpan(((DateTime)value).Ticks);
			}
			if (type == typeof(string))
			{
				return (T)(object)TimeSpan.Parse(value.ToString());
			}
		}
		return (T)Convert.ChangeType(value, typeFromHandle);
	}
}
