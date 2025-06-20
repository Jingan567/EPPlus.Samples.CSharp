using System;

namespace OfficeOpenXml.VBA;

public class ExcelVbaModuleCollection : ExcelVBACollectionBase<ExcelVBAModule>
{
	private ExcelVbaProject _project;

	internal ExcelVbaModuleCollection(ExcelVbaProject project)
	{
		_project = project;
	}

	internal void Add(ExcelVBAModule Item)
	{
		_list.Add(Item);
	}

	public ExcelVBAModule AddModule(string Name)
	{
		if (base[Name] != null)
		{
			throw new ArgumentException("Vba modulename already exist.");
		}
		ExcelVBAModule excelVBAModule = new ExcelVBAModule();
		excelVBAModule.Name = Name;
		excelVBAModule.Type = eModuleType.Module;
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_Name",
			Value = Name,
			DataType = eAttributeDataType.String
		});
		excelVBAModule.Type = eModuleType.Module;
		_list.Add(excelVBAModule);
		return excelVBAModule;
	}

	public ExcelVBAModule AddClass(string Name, bool Exposed)
	{
		ExcelVBAModule excelVBAModule = new ExcelVBAModule();
		excelVBAModule.Name = Name;
		excelVBAModule.Type = eModuleType.Class;
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_Name",
			Value = Name,
			DataType = eAttributeDataType.String
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_Base",
			Value = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}",
			DataType = eAttributeDataType.String
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_GlobalNameSpace",
			Value = "False",
			DataType = eAttributeDataType.NonString
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_Creatable",
			Value = "False",
			DataType = eAttributeDataType.NonString
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_PredeclaredId",
			Value = "False",
			DataType = eAttributeDataType.NonString
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_Exposed",
			Value = (Exposed ? "True" : "False"),
			DataType = eAttributeDataType.NonString
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_TemplateDerived",
			Value = "False",
			DataType = eAttributeDataType.NonString
		});
		excelVBAModule.Attributes._list.Add(new ExcelVbaModuleAttribute
		{
			Name = "VB_Customizable",
			Value = "False",
			DataType = eAttributeDataType.NonString
		});
		excelVBAModule.Private = !Exposed;
		_list.Add(excelVBAModule);
		return excelVBAModule;
	}
}
