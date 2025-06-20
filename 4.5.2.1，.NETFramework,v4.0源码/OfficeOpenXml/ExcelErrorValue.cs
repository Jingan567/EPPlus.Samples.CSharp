using System;
using System.Collections.Generic;

namespace OfficeOpenXml;

public class ExcelErrorValue
{
	public static class Values
	{
		public const string Div0 = "#DIV/0!";

		public const string NA = "#N/A";

		public const string Name = "#NAME?";

		public const string Null = "#NULL!";

		public const string Num = "#NUM!";

		public const string Ref = "#REF!";

		public const string Value = "#VALUE!";

		private static Dictionary<string, eErrorType> _values = new Dictionary<string, eErrorType>
		{
			{
				"#DIV/0!",
				eErrorType.Div0
			},
			{
				"#N/A",
				eErrorType.NA
			},
			{
				"#NAME?",
				eErrorType.Name
			},
			{
				"#NULL!",
				eErrorType.Null
			},
			{
				"#NUM!",
				eErrorType.Num
			},
			{
				"#REF!",
				eErrorType.Ref
			},
			{
				"#VALUE!",
				eErrorType.Value
			}
		};

		public static bool IsErrorValue(object candidate)
		{
			if (candidate == null || !(candidate is ExcelErrorValue))
			{
				return false;
			}
			string text = candidate.ToString();
			if (!string.IsNullOrEmpty(text))
			{
				return _values.ContainsKey(text);
			}
			return false;
		}

		public static bool StringIsErrorValue(string candidate)
		{
			if (!string.IsNullOrEmpty(candidate))
			{
				return _values.ContainsKey(candidate);
			}
			return false;
		}

		public static eErrorType ToErrorType(string val)
		{
			if (string.IsNullOrEmpty(val) || !_values.ContainsKey(val))
			{
				throw new ArgumentException("Invalid error code " + (val ?? "<empty>"));
			}
			return _values[val];
		}
	}

	public eErrorType Type { get; private set; }

	internal static ExcelErrorValue Create(eErrorType errorType)
	{
		return new ExcelErrorValue(errorType);
	}

	internal static ExcelErrorValue Parse(string val)
	{
		if (Values.StringIsErrorValue(val))
		{
			return new ExcelErrorValue(Values.ToErrorType(val));
		}
		if (string.IsNullOrEmpty(val))
		{
			throw new ArgumentNullException("val");
		}
		throw new ArgumentException("Not a valid error value: " + val);
	}

	private ExcelErrorValue(eErrorType type)
	{
		Type = type;
	}

	public override string ToString()
	{
		return Type switch
		{
			eErrorType.Div0 => "#DIV/0!", 
			eErrorType.NA => "#N/A", 
			eErrorType.Name => "#NAME?", 
			eErrorType.Null => "#NULL!", 
			eErrorType.Num => "#NUM!", 
			eErrorType.Ref => "#REF!", 
			eErrorType.Value => "#VALUE!", 
			_ => throw new ArgumentException("Invalid errortype"), 
		};
	}

	public static ExcelErrorValue operator +(object v1, ExcelErrorValue v2)
	{
		return v2;
	}

	public static ExcelErrorValue operator +(ExcelErrorValue v1, ExcelErrorValue v2)
	{
		return v1;
	}

	public override int GetHashCode()
	{
		return base.GetHashCode();
	}

	public override bool Equals(object obj)
	{
		if (!(obj is ExcelErrorValue))
		{
			return false;
		}
		return ((ExcelErrorValue)obj).ToString() == ToString();
	}
}
