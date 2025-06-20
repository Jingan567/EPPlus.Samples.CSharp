using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Exceptions;

public class ExcelErrorCodes
{
	private static readonly IEnumerable<string> Codes = new List<string> { Value.Code, Name.Code, NoValueAvaliable.Code };

	public string Code { get; private set; }

	public static ExcelErrorCodes Value => new ExcelErrorCodes("#VALUE!");

	public static ExcelErrorCodes Name => new ExcelErrorCodes("#NAME?");

	public static ExcelErrorCodes NoValueAvaliable => new ExcelErrorCodes("#N/A");

	private ExcelErrorCodes(string code)
	{
		Code = code;
	}

	public override int GetHashCode()
	{
		return Code.GetHashCode();
	}

	public override bool Equals(object obj)
	{
		if (obj is ExcelErrorCodes)
		{
			return ((ExcelErrorCodes)obj).Code.Equals(Code);
		}
		return false;
	}

	public static bool operator ==(ExcelErrorCodes c1, ExcelErrorCodes c2)
	{
		return c1.Code.Equals(c2.Code);
	}

	public static bool operator !=(ExcelErrorCodes c1, ExcelErrorCodes c2)
	{
		return !c1.Code.Equals(c2.Code);
	}

	public static bool IsErrorCode(object valueToTest)
	{
		if (valueToTest == null)
		{
			return false;
		}
		string candidate = valueToTest.ToString();
		if (Codes.FirstOrDefault((string x) => x == candidate) != null)
		{
			return true;
		}
		return false;
	}
}
