using System.Diagnostics;
using OfficeOpenXml.Packaging.Ionic.Zip;

namespace OfficeOpenXml.Packaging.Ionic;

internal abstract class SelectionCriterion
{
	internal virtual bool Verbose { get; set; }

	internal abstract bool Evaluate(string filename);

	[Conditional("SelectorTrace")]
	protected static void CriterionTrace(string format, params object[] args)
	{
	}

	internal abstract bool Evaluate(ZipEntry entry);
}
