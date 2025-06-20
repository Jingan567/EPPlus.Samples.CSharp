using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

public class WorkdayCalculatorResult
{
	public int NumberOfWorkdays { get; }

	public System.DateTime StartDate { get; }

	public System.DateTime EndDate { get; }

	public WorkdayCalculationDirection Direction { get; set; }

	public WorkdayCalculatorResult(int numberOfWorkdays, System.DateTime startDate, System.DateTime endDate, WorkdayCalculationDirection direction)
	{
		NumberOfWorkdays = numberOfWorkdays;
		StartDate = startDate;
		EndDate = endDate;
		Direction = direction;
	}
}
