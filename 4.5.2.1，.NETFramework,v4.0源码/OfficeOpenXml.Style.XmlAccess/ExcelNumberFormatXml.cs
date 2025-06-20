using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelNumberFormatXml : StyleXmlHelper
{
	internal enum eFormatType
	{
		Unknown,
		Number,
		DateTime
	}

	internal class ExcelFormatTranslator
	{
		private CultureInfo _ci;

		internal string NetTextFormat { get; private set; }

		internal string NetFormat { get; private set; }

		internal CultureInfo Culture
		{
			get
			{
				if (_ci == null)
				{
					return CultureInfo.CurrentCulture;
				}
				return _ci;
			}
			private set
			{
				_ci = value;
			}
		}

		internal eFormatType DataType { get; private set; }

		internal string NetTextFormatForWidth { get; private set; }

		internal string NetFormatForWidth { get; private set; }

		internal string FractionFormat { get; private set; }

		internal ExcelFormatTranslator(string format, int numFmtID)
		{
			if (numFmtID == 14)
			{
				NetFormat = (NetFormatForWidth = "d");
				NetTextFormat = (NetTextFormatForWidth = "");
				DataType = eFormatType.DateTime;
			}
			else if (format.Equals("general", StringComparison.OrdinalIgnoreCase))
			{
				NetFormat = (NetFormatForWidth = "0.#####");
				NetTextFormat = (NetTextFormatForWidth = "");
				DataType = eFormatType.Number;
			}
			else
			{
				ToNetFormat(format, forColWidth: false);
				ToNetFormat(format, forColWidth: true);
			}
		}

		private void ToNetFormat(string ExcelFormat, bool forColWidth)
		{
			DataType = eFormatType.Unknown;
			int num = 0;
			bool flag = false;
			bool flag2 = false;
			string text = "";
			bool flag3 = false;
			bool flag4 = false;
			bool flag5 = false;
			bool flag6 = false;
			string text2 = "";
			bool flag7 = ExcelFormat.Contains("AM/PM");
			List<int> list = new List<int>();
			StringBuilder stringBuilder = new StringBuilder();
			Culture = null;
			string text3 = "";
			string text4 = "";
			if (flag7)
			{
				ExcelFormat = Regex.Replace(ExcelFormat, "AM/PM", "");
				DataType = eFormatType.DateTime;
			}
			for (int i = 0; i < ExcelFormat.Length; i++)
			{
				char c = ExcelFormat[i];
				if (c == '"')
				{
					flag = !flag;
					continue;
				}
				if (flag6)
				{
					flag6 = false;
					continue;
				}
				if (flag && !flag2)
				{
					stringBuilder.Append(c);
					continue;
				}
				if (flag2)
				{
					if (c == ']')
					{
						flag2 = false;
						if (text[0] == '$')
						{
							string[] array = Regex.Split(text, "-");
							if (array[0].Length > 1)
							{
								stringBuilder.Append("\"" + array[0].Substring(1, array[0].Length - 1) + "\"");
							}
							if (array.Length <= 1)
							{
								continue;
							}
							if (array[1].Equals("f800", StringComparison.OrdinalIgnoreCase))
							{
								text2 = "D";
								continue;
							}
							if (array[1].Equals("f400", StringComparison.OrdinalIgnoreCase))
							{
								text2 = "T";
								continue;
							}
							int num2 = int.Parse(array[1], NumberStyles.HexNumber);
							try
							{
								Culture = CultureInfo.GetCultureInfo(num2 & 0xFFFF);
							}
							catch
							{
								Culture = null;
							}
						}
						else if (text[0] == 't')
						{
							stringBuilder.Append("hh");
						}
						else if (text[0] == 'h')
						{
							text2 = "hh";
						}
					}
					else
					{
						text += c;
					}
					continue;
				}
				if (flag5)
				{
					if (forColWidth)
					{
						stringBuilder.AppendFormat("\"{0}\"", c);
					}
					flag5 = false;
					continue;
				}
				if (c == ';')
				{
					num++;
					if (DataType == eFormatType.DateTime || num == 3)
					{
						if (DataType == eFormatType.DateTime)
						{
							SetDecimal(list, stringBuilder);
						}
						list = new List<int>();
						text3 = stringBuilder.ToString();
						stringBuilder = new StringBuilder();
					}
					else
					{
						stringBuilder.Append(c);
					}
					continue;
				}
				char c2 = c.ToString().ToLower(CultureInfo.InvariantCulture)[0];
				if (DataType == eFormatType.Unknown)
				{
					if (c == '0' || c == '#' || c == '.')
					{
						DataType = eFormatType.Number;
					}
					else if (c2 == 'y' || c2 == 'm' || c2 == 'd' || c2 == 'h' || c2 == 'm' || c2 == 's')
					{
						DataType = eFormatType.DateTime;
					}
				}
				if (flag3)
				{
					if (c == '.' || c == ',')
					{
						stringBuilder.Append('\\');
					}
					stringBuilder.Append(c);
					flag3 = false;
					continue;
				}
				switch (c)
				{
				case '[':
					text = "";
					flag2 = true;
					continue;
				case '\\':
					flag3 = true;
					continue;
				default:
					switch (c2)
					{
					case 'd':
					case 's':
						break;
					case 'h':
						if (flag7)
						{
							stringBuilder.Append('h');
						}
						else
						{
							stringBuilder.Append('H');
						}
						flag4 = true;
						continue;
					case 'm':
						if (flag4)
						{
							stringBuilder.Append('m');
						}
						else
						{
							stringBuilder.Append('M');
						}
						continue;
					default:
						switch (c)
						{
						case '_':
							flag5 = true;
							break;
						case '?':
							stringBuilder.Append(' ');
							break;
						case '/':
							if (DataType == eFormatType.Number)
							{
								_ = stringBuilder.Length;
								int num3 = i - 1;
								while (num3 >= 0 && (ExcelFormat[num3] == '?' || ExcelFormat[num3] == '#' || ExcelFormat[num3] == '0'))
								{
									num3--;
								}
								if (num3 > 0)
								{
									stringBuilder.Remove(stringBuilder.Length - (i - num3 - 1), i - num3 - 1);
								}
								int j;
								for (j = i + 1; j < ExcelFormat.Length && (ExcelFormat[j] == '?' || ExcelFormat[j] == '#' || (ExcelFormat[j] >= '0' && ExcelFormat[j] <= '9')); j++)
								{
								}
								i = j;
								if (FractionFormat != "")
								{
									FractionFormat = ExcelFormat.Substring(num3 + 1, j - num3 - 1);
								}
								stringBuilder.Append('?');
							}
							else
							{
								stringBuilder.Append('/');
							}
							break;
						case '*':
							flag6 = true;
							break;
						case '@':
							stringBuilder.Append("{0}");
							break;
						default:
							stringBuilder.Append(c);
							break;
						}
						continue;
					}
					break;
				case '#':
				case '%':
				case ',':
				case '.':
				case '0':
					break;
				}
				stringBuilder.Append(c);
				if (c == '.')
				{
					list.Add(stringBuilder.Length - 1);
				}
			}
			if (DataType == eFormatType.DateTime)
			{
				SetDecimal(list, stringBuilder);
			}
			if (flag7)
			{
				text3 += "tt";
			}
			if (text3 == "")
			{
				text3 = stringBuilder.ToString();
			}
			else
			{
				text4 = stringBuilder.ToString();
			}
			if (text2 != "")
			{
				text3 = text2;
			}
			if (forColWidth)
			{
				NetFormatForWidth = text3;
				NetTextFormatForWidth = text4;
			}
			else
			{
				NetFormat = text3;
				NetTextFormat = text4;
			}
			if (Culture == null)
			{
				Culture = CultureInfo.CurrentCulture;
			}
		}

		private static void SetDecimal(List<int> lstDec, StringBuilder sb)
		{
			if (lstDec.Count > 1)
			{
				for (int num = lstDec.Count - 1; num >= 0; num--)
				{
					sb.Insert(lstDec[num] + 1, '\'');
					sb.Insert(lstDec[num], '\'');
				}
			}
		}

		internal string FormatFraction(double d)
		{
			int num = (int)d;
			string[] array = FractionFormat.Split('/');
			if (!int.TryParse(array[1], out var result))
			{
				result = 0;
			}
			if (d == 0.0 || double.IsNaN(d))
			{
				if (array[0].Trim() == "" && array[1].Trim() == "")
				{
					return new string(' ', FractionFormat.Length);
				}
				return 0.ToString(array[0]) + "/" + 1.ToString(array[0]);
			}
			int length = array[1].Length;
			string text = ((d < 0.0) ? "-" : "");
			int num8;
			int num9;
			if (result == 0)
			{
				List<double> list = new List<double> { 1.0, 0.0 };
				List<double> list2 = new List<double> { 0.0, 1.0 };
				if (length < 1 && length > 12)
				{
					throw new ArgumentException("Number of digits out of range (1-12)");
				}
				int num2 = 0;
				for (int i = 0; i < length; i++)
				{
					num2 += 9 * (int)Math.Pow(10.0, i);
				}
				double num3 = 1.0 / (Math.Abs(d) - (double)num);
				double num4 = double.NaN;
				int index = 2;
				int num5 = 1;
				while (true)
				{
					num5++;
					double num6 = Math.Floor(num3);
					list.Add(num6 * list[num5 - 1] + list[num5 - 2]);
					if (list[num5] > (double)num2)
					{
						break;
					}
					list2.Add(num6 * list2[num5 - 1] + list2[num5 - 2]);
					double num7 = list[num5] / list2[num5];
					if (list2[num5] > (double)num2)
					{
						break;
					}
					index = num5;
					if (num7 == num4 || num7 == d)
					{
						break;
					}
					num4 = num7;
					num3 = 1.0 / (num3 - num6);
				}
				num8 = (int)list[index];
				num9 = (int)list2[index];
			}
			else
			{
				num8 = (int)Math.Round((d - (double)num) / (1.0 / (double)result), 0);
				num9 = result;
			}
			if (num8 == num9 || num8 == 0)
			{
				if (num8 == num9)
				{
					num++;
				}
				return text + num.ToString(NetFormat).Replace("?", new string(' ', FractionFormat.Length));
			}
			if (num == 0)
			{
				return text + FmtInt(num8, array[0]) + "/" + FmtInt(num9, array[1]);
			}
			return text + num.ToString(NetFormat).Replace("?", FmtInt(num8, array[0]) + "/" + FmtInt(num9, array[1]));
		}

		private string FmtInt(double value, string format)
		{
			string text = value.ToString("#");
			string text2 = "";
			if (text.Length < format.Length)
			{
				for (int num = format.Length - text.Length - 1; num >= 0; num--)
				{
					if (format[num] == '?')
					{
						text2 += " ";
					}
					else if (format[num] == ' ')
					{
						text2 += "0";
					}
				}
			}
			return text2 + text;
		}
	}

	private int _numFmtId;

	private const string fmtPath = "@formatCode";

	private string _format = string.Empty;

	private ExcelFormatTranslator _translator;

	public bool BuildIn { get; private set; }

	public int NumFmtId
	{
		get
		{
			return _numFmtId;
		}
		set
		{
			_numFmtId = value;
		}
	}

	internal override string Id => _format;

	public string Format
	{
		get
		{
			return _format;
		}
		set
		{
			_numFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
			_format = value;
		}
	}

	internal ExcelFormatTranslator FormatTranslator
	{
		get
		{
			if (_translator == null)
			{
				_translator = new ExcelFormatTranslator(Format, NumFmtId);
			}
			return _translator;
		}
	}

	internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
	}

	internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn)
		: base(nameSpaceManager)
	{
		BuildIn = buildIn;
	}

	internal ExcelNumberFormatXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		_numFmtId = GetXmlNodeInt("@numFmtId");
		_format = GetXmlNodeString("@formatCode");
	}

	internal string GetNewID(int NumFmtId, string Format)
	{
		if (NumFmtId < 0)
		{
			NumFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(Format);
		}
		return NumFmtId.ToString();
	}

	internal static void AddBuildIn(XmlNamespaceManager NameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats)
	{
		NumberFormats.Add("General", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 0,
			Format = "General"
		});
		NumberFormats.Add("0", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 1,
			Format = "0"
		});
		NumberFormats.Add("0.00", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 2,
			Format = "0.00"
		});
		NumberFormats.Add("#,##0", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 3,
			Format = "#,##0"
		});
		NumberFormats.Add("#,##0.00", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 4,
			Format = "#,##0.00"
		});
		NumberFormats.Add("0%", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 9,
			Format = "0%"
		});
		NumberFormats.Add("0.00%", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 10,
			Format = "0.00%"
		});
		NumberFormats.Add("0.00E+00", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 11,
			Format = "0.00E+00"
		});
		NumberFormats.Add("# ?/?", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 12,
			Format = "# ?/?"
		});
		NumberFormats.Add("# ??/??", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 13,
			Format = "# ??/??"
		});
		NumberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 14,
			Format = "mm-dd-yy"
		});
		NumberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 15,
			Format = "d-mmm-yy"
		});
		NumberFormats.Add("d-mmm", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 16,
			Format = "d-mmm"
		});
		NumberFormats.Add("mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 17,
			Format = "mmm-yy"
		});
		NumberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 18,
			Format = "h:mm AM/PM"
		});
		NumberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 19,
			Format = "h:mm:ss AM/PM"
		});
		NumberFormats.Add("h:mm", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 20,
			Format = "h:mm"
		});
		NumberFormats.Add("h:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 21,
			Format = "h:mm:ss"
		});
		NumberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 22,
			Format = "m/d/yy h:mm"
		});
		NumberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 37,
			Format = "#,##0 ;(#,##0)"
		});
		NumberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 38,
			Format = "#,##0 ;[Red](#,##0)"
		});
		NumberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 39,
			Format = "#,##0.00;(#,##0.00)"
		});
		NumberFormats.Add("#,##0.00;[Red](#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 40,
			Format = "#,##0.00;[Red](#,##0.00)"
		});
		NumberFormats.Add("mm:ss", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 45,
			Format = "mm:ss"
		});
		NumberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 46,
			Format = "[h]:mm:ss"
		});
		NumberFormats.Add("mmss.0", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 47,
			Format = "mmss.0"
		});
		NumberFormats.Add("##0.0", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 48,
			Format = "##0.0"
		});
		NumberFormats.Add("@", new ExcelNumberFormatXml(NameSpaceManager, buildIn: true)
		{
			NumFmtId = 49,
			Format = "@"
		});
		NumberFormats.NextId = 164;
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		SetXmlNodeString("@numFmtId", NumFmtId.ToString());
		SetXmlNodeString("@formatCode", Format);
		return base.TopNode;
	}
}
