using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class FormulaDependency
{
	private List<RangeAddress> _referencedBy = new List<RangeAddress>();

	private List<RangeAddress> _references = new List<RangeAddress>();

	public Guid ScopeId { get; private set; }

	public RangeAddress Address { get; private set; }

	public FormulaDependency(ParsingScope scope)
	{
		ScopeId = scope.ScopeId;
		Address = scope.Address;
	}

	public virtual void AddReferenceFrom(RangeAddress rangeAddress)
	{
		if (Address.CollidesWith(rangeAddress) || _references.Exists((RangeAddress x) => x.CollidesWith(rangeAddress)))
		{
			throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
		}
		_referencedBy.Add(rangeAddress);
	}

	public virtual void AddReferenceTo(RangeAddress rangeAddress)
	{
		if (Address.CollidesWith(rangeAddress) || _referencedBy.Exists((RangeAddress x) => x.CollidesWith(rangeAddress)))
		{
			throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
		}
		_references.Add(rangeAddress);
	}
}
