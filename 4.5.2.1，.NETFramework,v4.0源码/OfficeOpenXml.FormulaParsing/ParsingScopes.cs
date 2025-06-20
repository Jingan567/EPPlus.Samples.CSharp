using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing;

public class ParsingScopes
{
	private readonly IParsingLifetimeEventHandler _lifetimeEventHandler;

	private Stack<ParsingScope> _scopes = new Stack<ParsingScope>();

	public virtual ParsingScope Current
	{
		get
		{
			if (_scopes.Count() <= 0)
			{
				return null;
			}
			return _scopes.Peek();
		}
	}

	public ParsingScopes(IParsingLifetimeEventHandler lifetimeEventHandler)
	{
		_lifetimeEventHandler = lifetimeEventHandler;
	}

	public virtual ParsingScope NewScope(RangeAddress address)
	{
		ParsingScope parsingScope = ((_scopes.Count() <= 0) ? new ParsingScope(this, address) : new ParsingScope(this, _scopes.Peek(), address));
		_scopes.Push(parsingScope);
		return parsingScope;
	}

	public virtual void KillScope(ParsingScope parsingScope)
	{
		_scopes.Pop();
		if (_scopes.Count() == 0)
		{
			_lifetimeEventHandler.ParsingCompleted();
		}
	}
}
