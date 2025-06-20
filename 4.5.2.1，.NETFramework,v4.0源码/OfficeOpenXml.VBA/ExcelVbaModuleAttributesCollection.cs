using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.VBA;

public class ExcelVbaModuleAttributesCollection : ExcelVBACollectionBase<ExcelVbaModuleAttribute>
{
	internal string GetAttributeText()
	{
		StringBuilder stringBuilder = new StringBuilder();
		using (IEnumerator<ExcelVbaModuleAttribute> enumerator = GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ExcelVbaModuleAttribute current = enumerator.Current;
				stringBuilder.AppendFormat("Attribute {0} = {1}\r\n", current.Name, (current.DataType == eAttributeDataType.String) ? ("\"" + current.Value + "\"") : current.Value);
			}
		}
		return stringBuilder.ToString();
	}
}
