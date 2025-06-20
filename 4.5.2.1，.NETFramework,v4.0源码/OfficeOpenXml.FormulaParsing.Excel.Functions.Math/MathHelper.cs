using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public static class MathHelper
{
	public static double Sec(double x)
	{
		return 1.0 / System.Math.Cos(x);
	}

	public static double Cosec(double x)
	{
		return 1.0 / System.Math.Sin(x);
	}

	public static double Cotan(double x)
	{
		return 1.0 / System.Math.Tan(x);
	}

	public static double Arcsin(double x)
	{
		return System.Math.Atan(x / System.Math.Sqrt((0.0 - x) * x + 1.0));
	}

	public static double Arccos(double x)
	{
		return System.Math.Atan((0.0 - x) / System.Math.Sqrt((0.0 - x) * x + 1.0)) + 2.0 * System.Math.Atan(1.0);
	}

	public static double Arcsec(double x)
	{
		return 2.0 * System.Math.Atan(1.0) - System.Math.Atan((double)System.Math.Sign(x) / System.Math.Sqrt(x * x - 1.0));
	}

	public static double Arccosec(double x)
	{
		return System.Math.Atan((double)System.Math.Sign(x) / System.Math.Sqrt(x * x - 1.0));
	}

	public static double Arccotan(double x)
	{
		return 2.0 * System.Math.Atan(1.0) - System.Math.Atan(x);
	}

	public static double HSin(double x)
	{
		return (System.Math.Exp(x) - System.Math.Exp(0.0 - x)) / 2.0;
	}

	public static double HCos(double x)
	{
		return (System.Math.Exp(x) + System.Math.Exp(0.0 - x)) / 2.0;
	}

	public static double HTan(double x)
	{
		return (System.Math.Exp(x) - System.Math.Exp(0.0 - x)) / (System.Math.Exp(x) + System.Math.Exp(0.0 - x));
	}

	public static double HSec(double x)
	{
		return 2.0 / (System.Math.Exp(x) + System.Math.Exp(0.0 - x));
	}

	public static double HCosec(double x)
	{
		return 2.0 / (System.Math.Exp(x) - System.Math.Exp(0.0 - x));
	}

	public static double HCotan(double x)
	{
		return (System.Math.Exp(x) + System.Math.Exp(0.0 - x)) / (System.Math.Exp(x) - System.Math.Exp(0.0 - x));
	}

	public static double HArcsin(double x)
	{
		return System.Math.Log(x + System.Math.Sqrt(x * x + 1.0));
	}

	public static double HArccos(double x)
	{
		return System.Math.Log(x + System.Math.Sqrt(x * x - 1.0));
	}

	public static double HArctan(double x)
	{
		return System.Math.Log((1.0 + x) / (1.0 - x)) / 2.0;
	}

	public static double HArcsec(double x)
	{
		return System.Math.Log((System.Math.Sqrt((0.0 - x) * x + 1.0) + 1.0) / x);
	}

	public static double HArccosec(double x)
	{
		return System.Math.Log(((double)System.Math.Sign(x) * System.Math.Sqrt(x * x + 1.0) + 1.0) / x);
	}

	public static double HArccotan(double x)
	{
		return System.Math.Log((x + 1.0) / (x - 1.0)) / 2.0;
	}

	public static double LogN(double x, double n)
	{
		return System.Math.Log(x) / System.Math.Log(n);
	}
}
