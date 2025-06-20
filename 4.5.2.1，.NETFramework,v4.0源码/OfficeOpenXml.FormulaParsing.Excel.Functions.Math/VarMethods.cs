using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

internal static class VarMethods
{
	private static double Divide(double left, double right)
	{
		if (System.Math.Abs(right - 0.0) < double.Epsilon)
		{
			throw new ExcelErrorValueException(eErrorType.Div0);
		}
		return left / right;
	}

	public static double Var(IEnumerable<ExcelDoubleCellValue> args)
	{
		return Var(args.Select((Func<ExcelDoubleCellValue, double>)((ExcelDoubleCellValue x) => x)));
	}

	public static double Var(IEnumerable<double> args)
	{
		double avg = args.Select((double x) => x).Average();
		return Divide(args.Aggregate(0.0, (double total, double next) => total += System.Math.Pow(next - avg, 2.0)), args.Count() - 1);
	}

	public static double VarP(IEnumerable<ExcelDoubleCellValue> args)
	{
		return VarP(args.Select((Func<ExcelDoubleCellValue, double>)((ExcelDoubleCellValue x) => x)));
	}

	public static double VarP(IEnumerable<double> args)
	{
		double avg = args.Select((double x) => x).Average();
		return Divide(args.Aggregate(0.0, (double total, double next) => total += System.Math.Pow(next - avg, 2.0)), args.Count());
	}
}
