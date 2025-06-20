using System.Drawing;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingConstants
{
	internal class Errors
	{
		internal const string CommaSeparatedAddresses = "Multiple addresses may not be commaseparated, use space instead";

		internal const string InvalidCfruleObject = "The supplied item must inherit OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule";

		internal const string InvalidConditionalFormattingObject = "The supplied item must inherit OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormatting";

		internal const string InvalidPriority = "Invalid priority number. Must be bigger than zero";

		internal const string InvalidRemoveRuleOperation = "Invalid remove rule operation";

		internal const string MissingCfvoNode = "Missing 'cfvo' node in Conditional Formatting";

		internal const string MissingCfvoParentNode = "Missing 'cfvo' parent node in Conditional Formatting";

		internal const string MissingConditionalFormattingNode = "Missing 'conditionalFormatting' node in Conditional Formatting";

		internal const string MissingItemRuleList = "Missing item with address '{0}' in Conditional Formatting Rule List";

		internal const string MissingPriorityAttribute = "Missing 'priority' attribute in Conditional Formatting Rule";

		internal const string MissingRuleType = "Missing eExcelConditionalFormattingRuleType Type in Conditional Formatting";

		internal const string MissingSqrefAttribute = "Missing 'sqref' attribute in Conditional Formatting";

		internal const string MissingTypeAttribute = "Missing 'type' attribute in Conditional Formatting Rule";

		internal const string MissingWorksheetNode = "Missing 'worksheet' node";

		internal const string NonSupportedRuleType = "Non supported conditionalFormattingType: {0}";

		internal const string UnexistentCfvoTypeAttribute = "Unexistent eExcelConditionalFormattingValueObjectType attribute in Conditional Formatting";

		internal const string UnexistentOperatorTypeAttribute = "Unexistent eExcelConditionalFormattingOperatorType attribute in Conditional Formatting";

		internal const string UnexistentTimePeriodTypeAttribute = "Unexistent eExcelConditionalFormattingTimePeriodType attribute in Conditional Formatting";

		internal const string UnexpectedRuleTypeAttribute = "Unexpected eExcelConditionalFormattingRuleType attribute in Conditional Formatting Rule";

		internal const string UnexpectedRuleTypeName = "Unexpected eExcelConditionalFormattingRuleType TypeName in Conditional Formatting Rule";

		internal const string WrongNumberCfvoColorNodes = "Wrong number of 'cfvo'/'color' nodes in Conditional Formatting Rule";
	}

	internal class Nodes
	{
		internal const string Worksheet = "worksheet";

		internal const string ConditionalFormatting = "conditionalFormatting";

		internal const string CfRule = "cfRule";

		internal const string ColorScale = "colorScale";

		internal const string Cfvo = "cfvo";

		internal const string Color = "color";

		internal const string DataBar = "dataBar";

		internal const string IconSet = "iconSet";

		internal const string Formula = "formula";
	}

	internal class Attributes
	{
		internal const string AboveAverage = "aboveAverage";

		internal const string Bottom = "bottom";

		internal const string DxfId = "dxfId";

		internal const string EqualAverage = "equalAverage";

		internal const string IconSet = "iconSet";

		internal const string Operator = "operator";

		internal const string Percent = "percent";

		internal const string Priority = "priority";

		internal const string Rank = "rank";

		internal const string Reverse = "reverse";

		internal const string Rgb = "rgb";

		internal const string ShowValue = "showValue";

		internal const string Sqref = "sqref";

		internal const string StdDev = "stdDev";

		internal const string StopIfTrue = "stopIfTrue";

		internal const string Text = "text";

		internal const string Theme = "theme";

		internal const string TimePeriod = "timePeriod";

		internal const string Tint = "tint";

		internal const string Type = "type";

		internal const string Val = "val";

		internal const string Gte = "gte";
	}

	internal class Paths
	{
		internal const string Worksheet = "d:worksheet";

		internal const string ConditionalFormatting = "d:conditionalFormatting";

		internal const string CfRule = "d:cfRule";

		internal const string ColorScale = "d:colorScale";

		internal const string Cfvo = "d:cfvo";

		internal const string Color = "d:color";

		internal const string DataBar = "d:dataBar";

		internal const string IconSet = "d:iconSet";

		internal const string Formula = "d:formula";

		internal const string AboveAverageAttribute = "@aboveAverage";

		internal const string BottomAttribute = "@bottom";

		internal const string DxfIdAttribute = "@dxfId";

		internal const string EqualAverageAttribute = "@equalAverage";

		internal const string IconSetAttribute = "@iconSet";

		internal const string OperatorAttribute = "@operator";

		internal const string PercentAttribute = "@percent";

		internal const string PriorityAttribute = "@priority";

		internal const string RankAttribute = "@rank";

		internal const string ReverseAttribute = "@reverse";

		internal const string RgbAttribute = "@rgb";

		internal const string ShowValueAttribute = "@showValue";

		internal const string SqrefAttribute = "@sqref";

		internal const string StdDevAttribute = "@stdDev";

		internal const string StopIfTrueAttribute = "@stopIfTrue";

		internal const string TextAttribute = "@text";

		internal const string ThemeAttribute = "@theme";

		internal const string TimePeriodAttribute = "@timePeriod";

		internal const string TintAttribute = "@tint";

		internal const string TypeAttribute = "@type";

		internal const string ValAttribute = "@val";

		internal const string GteAttribute = "@gte";
	}

	internal class RuleType
	{
		internal const string AboveAverage = "aboveAverage";

		internal const string BeginsWith = "beginsWith";

		internal const string CellIs = "cellIs";

		internal const string ColorScale = "colorScale";

		internal const string ContainsBlanks = "containsBlanks";

		internal const string ContainsErrors = "containsErrors";

		internal const string ContainsText = "containsText";

		internal const string DataBar = "dataBar";

		internal const string DuplicateValues = "duplicateValues";

		internal const string EndsWith = "endsWith";

		internal const string Expression = "expression";

		internal const string IconSet = "iconSet";

		internal const string NotContainsBlanks = "notContainsBlanks";

		internal const string NotContainsErrors = "notContainsErrors";

		internal const string NotContainsText = "notContainsText";

		internal const string TimePeriod = "timePeriod";

		internal const string Top10 = "top10";

		internal const string UniqueValues = "uniqueValues";

		internal const string AboveOrEqualAverage = "aboveOrEqualAverage";

		internal const string AboveStdDev = "aboveStdDev";

		internal const string BelowAverage = "belowAverage";

		internal const string BelowOrEqualAverage = "belowOrEqualAverage";

		internal const string BelowStdDev = "belowStdDev";

		internal const string Between = "between";

		internal const string Bottom = "bottom";

		internal const string BottomPercent = "bottomPercent";

		internal const string Equal = "equal";

		internal const string GreaterThan = "greaterThan";

		internal const string GreaterThanOrEqual = "greaterThanOrEqual";

		internal const string IconSet3 = "iconSet3";

		internal const string IconSet4 = "iconSet4";

		internal const string IconSet5 = "iconSet5";

		internal const string Last7Days = "last7Days";

		internal const string LastMonth = "lastMonth";

		internal const string LastWeek = "lastWeek";

		internal const string LessThan = "lessThan";

		internal const string LessThanOrEqual = "lessThanOrEqual";

		internal const string NextMonth = "nextMonth";

		internal const string NextWeek = "nextWeek";

		internal const string NotBetween = "notBetween";

		internal const string NotEqual = "notEqual";

		internal const string ThisMonth = "thisMonth";

		internal const string ThisWeek = "thisWeek";

		internal const string ThreeColorScale = "threeColorScale";

		internal const string Today = "today";

		internal const string Tomorrow = "tomorrow";

		internal const string Top = "top";

		internal const string TopPercent = "topPercent";

		internal const string TwoColorScale = "twoColorScale";

		internal const string Yesterday = "yesterday";
	}

	internal class CfvoType
	{
		internal const string Min = "min";

		internal const string Max = "max";

		internal const string Num = "num";

		internal const string Formula = "formula";

		internal const string Percent = "percent";

		internal const string Percentile = "percentile";
	}

	internal class Operators
	{
		internal const string BeginsWith = "beginsWith";

		internal const string Between = "between";

		internal const string ContainsText = "containsText";

		internal const string EndsWith = "endsWith";

		internal const string Equal = "equal";

		internal const string GreaterThan = "greaterThan";

		internal const string GreaterThanOrEqual = "greaterThanOrEqual";

		internal const string LessThan = "lessThan";

		internal const string LessThanOrEqual = "lessThanOrEqual";

		internal const string NotBetween = "notBetween";

		internal const string NotContains = "notContains";

		internal const string NotEqual = "notEqual";
	}

	internal class TimePeriods
	{
		internal const string Last7Days = "last7Days";

		internal const string LastMonth = "lastMonth";

		internal const string LastWeek = "lastWeek";

		internal const string NextMonth = "nextMonth";

		internal const string NextWeek = "nextWeek";

		internal const string ThisMonth = "thisMonth";

		internal const string ThisWeek = "thisWeek";

		internal const string Today = "today";

		internal const string Tomorrow = "tomorrow";

		internal const string Yesterday = "yesterday";
	}

	internal class Colors
	{
		internal static readonly Color CfvoLowValue = Color.FromArgb(255, 248, 105, 107);

		internal static readonly Color CfvoMiddleValue = Color.FromArgb(255, 255, 235, 132);

		internal static readonly Color CfvoHighValue = Color.FromArgb(255, 99, 190, 123);
	}
}
