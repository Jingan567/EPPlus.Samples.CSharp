using System;
using System.Globalization;

namespace OfficeOpenXml.DataValidation;

public class ExcelTime
{
	private readonly decimal SecondsPerDay = 86400m;

	private readonly decimal SecondsPerHour = 3600m;

	private readonly decimal SecondsPerMinute = 60m;

	public const int NumberOfDecimals = 15;

	private int _hour;

	private int _minute;

	private int? _second;

	public int Hour
	{
		get
		{
			return _hour;
		}
		set
		{
			if (value < 0)
			{
				throw new InvalidOperationException("Value for hour cannot be negative");
			}
			if (value > 23)
			{
				throw new InvalidOperationException("Value for hour cannot be greater than 23");
			}
			_hour = value;
			OnTimeChanged();
		}
	}

	public int Minute
	{
		get
		{
			return _minute;
		}
		set
		{
			if (value < 0)
			{
				throw new InvalidOperationException("Value for minute cannot be negative");
			}
			if (value > 59)
			{
				throw new InvalidOperationException("Value for minute cannot be greater than 59");
			}
			_minute = value;
			OnTimeChanged();
		}
	}

	public int? Second
	{
		get
		{
			return _second;
		}
		set
		{
			if (value < 0)
			{
				throw new InvalidOperationException("Value for second cannot be negative");
			}
			if (value > 59)
			{
				throw new InvalidOperationException("Value for second cannot be greater than 59");
			}
			_second = value;
			OnTimeChanged();
		}
	}

	private event EventHandler _timeChanged;

	internal event EventHandler TimeChanged
	{
		add
		{
			_timeChanged += value;
		}
		remove
		{
			_timeChanged -= value;
		}
	}

	public ExcelTime()
	{
	}

	public ExcelTime(decimal value)
	{
		if (value < 0m)
		{
			throw new ArgumentException("Value cannot be less than 0");
		}
		if (value >= 1m)
		{
			throw new ArgumentException("Value cannot be greater or equal to 1");
		}
		Init(value);
	}

	private void Init(decimal value)
	{
		decimal num = value * SecondsPerDay;
		decimal num2 = Math.Floor(num / SecondsPerHour);
		Hour = (int)num2;
		decimal num3 = Math.Floor((num - num2 * SecondsPerHour) / SecondsPerMinute);
		Minute = (int)num3;
		decimal num4 = Math.Round(num - num2 * SecondsPerHour - num3 * SecondsPerMinute, MidpointRounding.AwayFromZero);
		SetSecond((int)num4);
	}

	private void SetSecond(int value)
	{
		if (value == 60)
		{
			Second = 0;
			int minute = Minute + 1;
			SetMinute(minute);
		}
		else
		{
			Second = value;
		}
	}

	private void SetMinute(int value)
	{
		if (value == 60)
		{
			Minute = 0;
			int hour = Hour + 1;
			SetHour(hour);
		}
		else
		{
			Minute = value;
		}
	}

	private void SetHour(int value)
	{
		if (value == 24)
		{
			Hour = 0;
		}
	}

	private void OnTimeChanged()
	{
		if (this._timeChanged != null)
		{
			this._timeChanged(this, EventArgs.Empty);
		}
	}

	private decimal Round(decimal value)
	{
		return Math.Round(value, 15);
	}

	private decimal ToSeconds()
	{
		return (decimal)Hour * SecondsPerHour + (decimal)Minute * SecondsPerMinute + (decimal)(Second ?? 0);
	}

	public decimal ToExcelTime()
	{
		decimal num = ToSeconds();
		return Round(num / SecondsPerDay);
	}

	public string ToExcelString()
	{
		return ToExcelTime().ToString(CultureInfo.InvariantCulture);
	}

	public override string ToString()
	{
		int num = Second ?? 0;
		return string.Format("{0}:{1}:{2}", (Hour < 10) ? ("0" + Hour) : Hour.ToString(), (Minute < 10) ? ("0" + Minute) : Minute.ToString(), (num < 10) ? ("0" + num) : num.ToString());
	}
}
