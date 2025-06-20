using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationCollection : XmlHelper, IEnumerable<IExcelDataValidation>, IEnumerable
{
	private List<IExcelDataValidation> _validations = new List<IExcelDataValidation>();

	private ExcelWorksheet _worksheet;

	private const string DataValidationPath = "//d:dataValidations";

	private readonly string DataValidationItemsPath = string.Format("{0}/d:dataValidation", "//d:dataValidations");

	public int Count => _validations.Count;

	public IExcelDataValidation this[int index]
	{
		get
		{
			return _validations[index];
		}
		set
		{
			_validations[index] = value;
		}
	}

	public IExcelDataValidation this[string address]
	{
		get
		{
			ExcelAddress searchedAddress = new ExcelAddress(address);
			return _validations.Find((IExcelDataValidation x) => x.Address.Collide(searchedAddress) != ExcelAddressBase.eAddressCollition.No);
		}
	}

	internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
		: base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
	{
		Require.Argument(worksheet).IsNotNull("worksheet");
		_worksheet = worksheet;
		base.SchemaNodeOrder = worksheet.SchemaNodeOrder;
		XmlNodeList xmlNodeList = worksheet.WorksheetXml.SelectNodes(DataValidationItemsPath, worksheet.NameSpaceManager);
		if (xmlNodeList != null && xmlNodeList.Count > 0)
		{
			foreach (XmlNode item in xmlNodeList)
			{
				if (item.Attributes["sqref"] != null)
				{
					string value = item.Attributes["sqref"].Value;
					ExcelDataValidationType bySchemaName = ExcelDataValidationType.GetBySchemaName((item.Attributes["type"] != null) ? item.Attributes["type"].Value : "");
					_validations.Add(ExcelDataValidationFactory.Create(bySchemaName, worksheet, value, item));
				}
			}
		}
		if (_validations.Count > 0)
		{
			OnValidationCountChanged();
		}
	}

	private void EnsureRootElementExists()
	{
		if (_worksheet.WorksheetXml.SelectSingleNode("//d:dataValidations", _worksheet.NameSpaceManager) == null)
		{
			CreateNode("//d:dataValidations".TrimStart('/'));
		}
	}

	private void OnValidationCountChanged()
	{
		XmlNode rootNode = GetRootNode();
		if (_validations.Count == 0)
		{
			if (rootNode != null)
			{
				_worksheet.WorksheetXml.DocumentElement.RemoveChild(rootNode);
			}
			_worksheet.ClearValidations();
		}
		else
		{
			if (_worksheet.WorksheetXml.DocumentElement.SelectSingleNode("//d:dataValidations[@count]", _worksheet.NameSpaceManager) == null)
			{
				rootNode.Attributes.Append(_worksheet.WorksheetXml.CreateAttribute("count"));
			}
			rootNode.Attributes["count"].Value = _validations.Count.ToString(CultureInfo.InvariantCulture);
		}
	}

	private XmlNode GetRootNode()
	{
		EnsureRootElementExists();
		base.TopNode = _worksheet.WorksheetXml.SelectSingleNode("//d:dataValidations", _worksheet.NameSpaceManager);
		return base.TopNode;
	}

	private void ValidateAddress(string address, IExcelDataValidation validatingValidation)
	{
		Require.Argument(address).IsNotNullOrEmpty("address");
		ExcelAddress address2 = new ExcelAddress(address);
		if (_validations.Count <= 0)
		{
			return;
		}
		foreach (IExcelDataValidation validation in _validations)
		{
			if ((validatingValidation == null || validatingValidation != validation) && validation.Address.Collide(address2) != 0)
			{
				throw new InvalidOperationException($"The address ({address}) collides with an existing validation ({validation.Address.Address})");
			}
		}
	}

	private void ValidateAddress(string address)
	{
		ValidateAddress(address, null);
	}

	internal void ValidateAll()
	{
		foreach (IExcelDataValidation validation in _validations)
		{
			validation.Validate();
			ValidateAddress(validation.Address.Address, validation);
		}
	}

	public IExcelDataValidationAny AddAnyValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationAny excelDataValidationAny = new ExcelDataValidationAny(_worksheet, address, ExcelDataValidationType.Any);
		_validations.Add(excelDataValidationAny);
		OnValidationCountChanged();
		return excelDataValidationAny;
	}

	public IExcelDataValidationInt AddIntegerValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationInt excelDataValidationInt = new ExcelDataValidationInt(_worksheet, address, ExcelDataValidationType.Whole);
		_validations.Add(excelDataValidationInt);
		OnValidationCountChanged();
		return excelDataValidationInt;
	}

	public IExcelDataValidationDecimal AddDecimalValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationDecimal excelDataValidationDecimal = new ExcelDataValidationDecimal(_worksheet, address, ExcelDataValidationType.Decimal);
		_validations.Add(excelDataValidationDecimal);
		OnValidationCountChanged();
		return excelDataValidationDecimal;
	}

	public IExcelDataValidationList AddListValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationList excelDataValidationList = new ExcelDataValidationList(_worksheet, address, ExcelDataValidationType.List);
		_validations.Add(excelDataValidationList);
		OnValidationCountChanged();
		return excelDataValidationList;
	}

	public IExcelDataValidationInt AddTextLengthValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationInt excelDataValidationInt = new ExcelDataValidationInt(_worksheet, address, ExcelDataValidationType.TextLength);
		_validations.Add(excelDataValidationInt);
		OnValidationCountChanged();
		return excelDataValidationInt;
	}

	public IExcelDataValidationDateTime AddDateTimeValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationDateTime excelDataValidationDateTime = new ExcelDataValidationDateTime(_worksheet, address, ExcelDataValidationType.DateTime);
		_validations.Add(excelDataValidationDateTime);
		OnValidationCountChanged();
		return excelDataValidationDateTime;
	}

	public IExcelDataValidationTime AddTimeValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationTime excelDataValidationTime = new ExcelDataValidationTime(_worksheet, address, ExcelDataValidationType.Time);
		_validations.Add(excelDataValidationTime);
		OnValidationCountChanged();
		return excelDataValidationTime;
	}

	public IExcelDataValidationCustom AddCustomValidation(string address)
	{
		ValidateAddress(address);
		EnsureRootElementExists();
		ExcelDataValidationCustom excelDataValidationCustom = new ExcelDataValidationCustom(_worksheet, address, ExcelDataValidationType.Custom);
		_validations.Add(excelDataValidationCustom);
		OnValidationCountChanged();
		return excelDataValidationCustom;
	}

	public bool Remove(IExcelDataValidation item)
	{
		if (!(item is ExcelDataValidation))
		{
			throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
		}
		Require.Argument(item).IsNotNull("item");
		_worksheet.WorksheetXml.DocumentElement.SelectSingleNode("//d:dataValidations".TrimStart('/'), base.NameSpaceManager)?.RemoveChild(((ExcelDataValidation)item).TopNode);
		bool num = _validations.Remove(item);
		if (num)
		{
			OnValidationCountChanged();
		}
		return num;
	}

	public IEnumerable<IExcelDataValidation> FindAll(Predicate<IExcelDataValidation> match)
	{
		return _validations.FindAll(match);
	}

	public IExcelDataValidation Find(Predicate<IExcelDataValidation> match)
	{
		return _validations.Find(match);
	}

	public void Clear()
	{
		DeleteAllNode(DataValidationItemsPath.TrimStart('/'));
		_validations.Clear();
	}

	public void RemoveAll(Predicate<IExcelDataValidation> match)
	{
		foreach (IExcelDataValidation item in _validations.FindAll(match))
		{
			if (!(item is ExcelDataValidation))
			{
				throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
			}
			base.TopNode.RemoveChild(((ExcelDataValidation)item).TopNode);
		}
		_validations.RemoveAll(match);
		OnValidationCountChanged();
	}

	IEnumerator<IExcelDataValidation> IEnumerable<IExcelDataValidation>.GetEnumerator()
	{
		return _validations.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _validations.GetEnumerator();
	}
}
