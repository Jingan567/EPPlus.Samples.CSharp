using System;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationType
{
	private static ExcelDataValidationType _any;

	private static ExcelDataValidationType _whole;

	private static ExcelDataValidationType _list;

	private static ExcelDataValidationType _decimal;

	private static ExcelDataValidationType _textLength;

	private static ExcelDataValidationType _dateTime;

	private static ExcelDataValidationType _time;

	private static ExcelDataValidationType _custom;

	public eDataValidationType Type { get; private set; }

	internal string SchemaName { get; private set; }

	internal bool AllowOperator { get; private set; }

	public static ExcelDataValidationType Any
	{
		get
		{
			if (_any == null)
			{
				_any = new ExcelDataValidationType(eDataValidationType.Any, allowOperator: false, "");
			}
			return _any;
		}
	}

	public static ExcelDataValidationType Whole
	{
		get
		{
			if (_whole == null)
			{
				_whole = new ExcelDataValidationType(eDataValidationType.Whole, allowOperator: true, "whole");
			}
			return _whole;
		}
	}

	public static ExcelDataValidationType List
	{
		get
		{
			if (_list == null)
			{
				_list = new ExcelDataValidationType(eDataValidationType.List, allowOperator: false, "list");
			}
			return _list;
		}
	}

	public static ExcelDataValidationType Decimal
	{
		get
		{
			if (_decimal == null)
			{
				_decimal = new ExcelDataValidationType(eDataValidationType.Decimal, allowOperator: true, "decimal");
			}
			return _decimal;
		}
	}

	public static ExcelDataValidationType TextLength
	{
		get
		{
			if (_textLength == null)
			{
				_textLength = new ExcelDataValidationType(eDataValidationType.TextLength, allowOperator: true, "textLength");
			}
			return _textLength;
		}
	}

	public static ExcelDataValidationType DateTime
	{
		get
		{
			if (_dateTime == null)
			{
				_dateTime = new ExcelDataValidationType(eDataValidationType.DateTime, allowOperator: true, "date");
			}
			return _dateTime;
		}
	}

	public static ExcelDataValidationType Time
	{
		get
		{
			if (_time == null)
			{
				_time = new ExcelDataValidationType(eDataValidationType.Time, allowOperator: true, "time");
			}
			return _time;
		}
	}

	public static ExcelDataValidationType Custom
	{
		get
		{
			if (_custom == null)
			{
				_custom = new ExcelDataValidationType(eDataValidationType.Custom, allowOperator: true, "custom");
			}
			return _custom;
		}
	}

	private ExcelDataValidationType(eDataValidationType validationType, bool allowOperator, string schemaName)
	{
		Type = validationType;
		AllowOperator = allowOperator;
		SchemaName = schemaName;
	}

	internal static ExcelDataValidationType GetByValidationType(eDataValidationType type)
	{
		return type switch
		{
			eDataValidationType.Any => Any, 
			eDataValidationType.Whole => Whole, 
			eDataValidationType.List => List, 
			eDataValidationType.Decimal => Decimal, 
			eDataValidationType.TextLength => TextLength, 
			eDataValidationType.DateTime => DateTime, 
			eDataValidationType.Time => Time, 
			eDataValidationType.Custom => Custom, 
			_ => throw new InvalidOperationException("Non supported Validationtype : " + type), 
		};
	}

	internal static ExcelDataValidationType GetBySchemaName(string schemaName)
	{
		return schemaName switch
		{
			"" => Any, 
			"whole" => Whole, 
			"decimal" => Decimal, 
			"list" => List, 
			"textLength" => TextLength, 
			"date" => DateTime, 
			"time" => Time, 
			"custom" => Custom, 
			_ => throw new ArgumentException("Invalid schemaname: " + schemaName), 
		};
	}

	public override bool Equals(object obj)
	{
		if (!(obj is ExcelDataValidationType))
		{
			return false;
		}
		return ((ExcelDataValidationType)obj).Type == Type;
	}

	public override int GetHashCode()
	{
		return base.GetHashCode();
	}
}
