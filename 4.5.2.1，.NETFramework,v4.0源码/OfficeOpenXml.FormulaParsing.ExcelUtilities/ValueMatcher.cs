using System;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class ValueMatcher
{
	public const int IncompatibleOperands = -2;

	public virtual int IsMatch(object o1, object o2)
	{
		if (o1 != null && o2 == null)
		{
			return 1;
		}
		if (o1 == null && o2 != null)
		{
			return -1;
		}
		if (o1 == null && o2 == null)
		{
			return 0;
		}
		o1 = CheckGetRange(o1);
		o2 = CheckGetRange(o2);
		if (o1 is string && o2 is string)
		{
			return CompareStringToString(o1.ToString().ToLower(), o2.ToString().ToLower());
		}
		if (o1.GetType() == typeof(string))
		{
			return CompareStringToObject(o1.ToString(), o2);
		}
		if (o2.GetType() == typeof(string))
		{
			return CompareObjectToString(o1, o2.ToString());
		}
		return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
	}

	private static object CheckGetRange(object v)
	{
		if (v is ExcelDataProvider.IRangeInfo)
		{
			ExcelDataProvider.IRangeInfo obj = (ExcelDataProvider.IRangeInfo)v;
			if (obj.GetNCells() > 1)
			{
				v = ExcelErrorValue.Create(eErrorType.NA);
			}
			v = obj.GetOffset(0, 0);
		}
		else if (v is ExcelDataProvider.INameInfo)
		{
			v = CheckGetRange((ExcelDataProvider.INameInfo)v);
		}
		return v;
	}

	protected virtual int CompareStringToString(string s1, string s2)
	{
		return s1.CompareTo(s2);
	}

	protected virtual int CompareStringToObject(string o1, object o2)
	{
		if (double.TryParse(o1, out var result))
		{
			return result.CompareTo(Convert.ToDouble(o2));
		}
		if (bool.TryParse(o1, out var result2))
		{
			return result2.CompareTo(Convert.ToBoolean(o2));
		}
		if (DateTime.TryParse(o1, out var result3))
		{
			return result3.CompareTo(Convert.ToDateTime(o2));
		}
		return -2;
	}

	protected virtual int CompareObjectToString(object o1, string o2)
	{
		if (double.TryParse(o2, out var result))
		{
			return Convert.ToDouble(o1).CompareTo(result);
		}
		return -2;
	}
}
