using System;
using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml.FormulaParsing.Utilities;

public static class ExtensionMethods
{
	public static void IsNotNullOrEmpty(this ArgumentInfo<string> val)
	{
		if (string.IsNullOrEmpty(val.Value))
		{
			throw new ArgumentException(val.Name + " cannot be null or empty");
		}
	}

	public static void IsNotNull<T>(this ArgumentInfo<T> val) where T : class
	{
		if (val.Value == null)
		{
			throw new ArgumentNullException(val.Name);
		}
	}

	public static bool IsNumeric(this object obj)
	{
		if (obj == null)
		{
			return false;
		}
		if (!TypeCompat.IsPrimitive(obj) && !(obj is double) && !(obj is decimal) && !(obj is DateTime))
		{
			return obj is TimeSpan;
		}
		return true;
	}
}
