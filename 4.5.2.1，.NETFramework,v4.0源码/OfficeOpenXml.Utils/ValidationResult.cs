using System;

namespace OfficeOpenXml.Utils;

public class ValidationResult : IValidationResult
{
	private bool _result;

	private string _errorMessage;

	public ValidationResult(bool result)
		: this(result, null)
	{
	}

	public ValidationResult(bool result, string errorMessage)
	{
		_result = result;
		_errorMessage = errorMessage;
	}

	private void Throw()
	{
		if (string.IsNullOrEmpty(_errorMessage))
		{
			throw new InvalidOperationException();
		}
		throw new InvalidOperationException(_errorMessage);
	}

	void IValidationResult.IsTrue()
	{
		if (!_result)
		{
			Throw();
		}
	}

	void IValidationResult.IsFalse()
	{
		if (_result)
		{
			Throw();
		}
	}
}
