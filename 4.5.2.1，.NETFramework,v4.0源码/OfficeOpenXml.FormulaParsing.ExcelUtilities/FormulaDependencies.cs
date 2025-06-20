using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class FormulaDependencies
{
	private readonly FormulaDependencyFactory _formulaDependencyFactory;

	private readonly Dictionary<string, FormulaDependency> _dependencies = new Dictionary<string, FormulaDependency>();

	public IEnumerable<KeyValuePair<string, FormulaDependency>> Dependencies => _dependencies;

	public FormulaDependencies()
		: this(new FormulaDependencyFactory())
	{
	}

	public FormulaDependencies(FormulaDependencyFactory formulaDependencyFactory)
	{
		_formulaDependencyFactory = formulaDependencyFactory;
	}

	public void AddFormulaScope(ParsingScope parsingScope)
	{
	}

	public void Clear()
	{
		_dependencies.Clear();
	}
}
