using System.ComponentModel;

namespace OfficeOpenXml.Packaging.Ionic;

internal enum ComparisonOperator
{
	[Description(">")]
	GreaterThan,
	[Description(">=")]
	GreaterThanOrEqualTo,
	[Description("<")]
	LesserThan,
	[Description("<=")]
	LesserThanOrEqualTo,
	[Description("=")]
	EqualTo,
	[Description("!=")]
	NotEqualTo
}
