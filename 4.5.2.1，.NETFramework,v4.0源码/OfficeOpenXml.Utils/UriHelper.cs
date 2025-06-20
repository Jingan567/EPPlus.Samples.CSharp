using System;

namespace OfficeOpenXml.Utils;

internal class UriHelper
{
	internal static Uri ResolvePartUri(Uri sourceUri, Uri targetUri)
	{
		if (targetUri.OriginalString.StartsWith("/") || targetUri.OriginalString.Contains("://"))
		{
			return targetUri;
		}
		string[] array = sourceUri.OriginalString.Split('/');
		string[] array2 = targetUri.OriginalString.Split('/');
		int num = array2.Length - 1;
		int num2 = ((!sourceUri.OriginalString.EndsWith("/")) ? (array.Length - 2) : (array.Length - 1));
		string text = array2[num--];
		while (num >= 0 && !(array2[num] == "."))
		{
			if (array2[num] == "..")
			{
				num2--;
				num--;
			}
			else
			{
				text = array2[num--] + "/" + text;
			}
		}
		if (num2 >= 0)
		{
			for (int num3 = num2; num3 >= 0; num3--)
			{
				text = array[num3] + "/" + text;
			}
		}
		return new Uri(text, UriKind.RelativeOrAbsolute);
	}

	internal static Uri GetRelativeUri(Uri WorksheetUri, Uri uri)
	{
		string[] array = WorksheetUri.OriginalString.Split('/');
		string[] array2 = uri.OriginalString.Split('/');
		int num = ((!WorksheetUri.OriginalString.EndsWith("/")) ? (array.Length - 1) : array.Length);
		int i;
		for (i = 0; i < num && i < array2.Length && array[i] == array2[i]; i++)
		{
		}
		string text = "";
		for (int j = i; j < num; j++)
		{
			text += "../";
		}
		string text2 = "";
		for (int k = i; k < array2.Length; k++)
		{
			text2 = text2 + ((text2 == "") ? "" : "/") + array2[k];
		}
		return new Uri(text + text2, UriKind.Relative);
	}
}
