using System;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.VBA;

public class ExcelVbaProtection
{
	private ExcelVbaProject _project;

	public bool UserProtected { get; internal set; }

	public bool HostProtected { get; internal set; }

	public bool VbeProtected { get; internal set; }

	public bool VisibilityState { get; internal set; }

	internal byte[] PasswordHash { get; set; }

	internal byte[] PasswordKey { get; set; }

	internal ExcelVbaProtection(ExcelVbaProject project)
	{
		_project = project;
		VisibilityState = true;
	}

	public void SetPassword(string Password)
	{
		if (string.IsNullOrEmpty(Password))
		{
			PasswordHash = null;
			PasswordKey = null;
			VbeProtected = false;
			HostProtected = false;
			UserProtected = false;
			VisibilityState = true;
			_project.ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
		}
		else
		{
			PasswordKey = new byte[4];
			RandomNumberGenerator.Create().GetBytes(PasswordKey);
			byte[] array = new byte[Password.Length + 4];
			Array.Copy(Encoding.GetEncoding(_project.CodePage).GetBytes(Password), array, Password.Length);
			VbeProtected = true;
			VisibilityState = false;
			Array.Copy(PasswordKey, 0, array, array.Length - 4, 4);
			SHA1 sHA = SHA1.Create();
			PasswordHash = sHA.ComputeHash(array);
			_project.ProjectID = "{00000000-0000-0000-0000-000000000000}";
		}
	}
}
