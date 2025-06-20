using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation;

internal static class ExcelDataValidationFactory
{
	public static ExcelDataValidation Create(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode)
	{
		Require.Argument(type).IsNotNull("validationType");
		switch (type.Type)
		{
		case eDataValidationType.Any:
			return new ExcelDataValidationAny(worksheet, address, type, itemElementNode);
		case eDataValidationType.Whole:
		case eDataValidationType.TextLength:
			return new ExcelDataValidationInt(worksheet, address, type, itemElementNode);
		case eDataValidationType.Decimal:
			return new ExcelDataValidationDecimal(worksheet, address, type, itemElementNode);
		case eDataValidationType.List:
			return new ExcelDataValidationList(worksheet, address, type, itemElementNode);
		case eDataValidationType.DateTime:
			return new ExcelDataValidationDateTime(worksheet, address, type, itemElementNode);
		case eDataValidationType.Time:
			return new ExcelDataValidationTime(worksheet, address, type, itemElementNode);
		case eDataValidationType.Custom:
			return new ExcelDataValidationCustom(worksheet, address, type, itemElementNode);
		default:
			throw new InvalidOperationException("Non supported validationtype: " + type.Type);
		}
	}
}
