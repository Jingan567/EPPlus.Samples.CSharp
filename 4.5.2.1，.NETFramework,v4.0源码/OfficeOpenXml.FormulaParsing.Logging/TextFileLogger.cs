using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Logging;

internal class TextFileLogger : IFormulaParserLogger, IDisposable
{
	private StreamWriter _sw;

	private const string Separator = "=================================";

	private int _count;

	private DateTime _startTime = DateTime.Now;

	private Dictionary<string, int> _funcs = new Dictionary<string, int>();

	private Dictionary<string, long> _funcPerformance = new Dictionary<string, long>();

	internal TextFileLogger(FileInfo fileInfo)
	{
		_sw = new StreamWriter(fileInfo.FullName);
	}

	private void WriteSeparatorAndTimeStamp()
	{
		_sw.WriteLine("=================================");
		_sw.WriteLine("Timestamp: {0}", DateTime.Now);
		_sw.WriteLine();
	}

	private void WriteAddressInfo(ParsingContext context)
	{
		if (context.Scopes.Current != null && context.Scopes.Current.Address != null)
		{
			_sw.WriteLine("Worksheet: {0}", context.Scopes.Current.Address.Worksheet ?? "<not specified>");
			_sw.WriteLine("Address: {0}", context.Scopes.Current.Address.Address ?? "<not available>");
		}
	}

	public void Log(ParsingContext context, Exception ex)
	{
		WriteSeparatorAndTimeStamp();
		WriteAddressInfo(context);
		_sw.WriteLine(ex);
		_sw.WriteLine();
	}

	public void Log(ParsingContext context, string message)
	{
		WriteSeparatorAndTimeStamp();
		WriteAddressInfo(context);
		_sw.WriteLine(message);
		_sw.WriteLine();
	}

	public void Log(string message)
	{
		WriteSeparatorAndTimeStamp();
		_sw.WriteLine(message);
		_sw.WriteLine();
	}

	public void LogCellCounted()
	{
		_count++;
		if (_count % 500 != 0)
		{
			return;
		}
		_sw.WriteLine("=================================");
		TimeSpan timeSpan = DateTime.Now.Subtract(_startTime);
		_sw.WriteLine("{0} cells parsed, time {1} seconds", _count, timeSpan.TotalSeconds);
		foreach (string item in _funcs.Keys.OrderByDescending((string x) => _funcs[x]).ToList())
		{
			_sw.Write(item + "  - " + _funcs[item]);
			if (_funcPerformance.ContainsKey(item))
			{
				_sw.Write(" - avg: " + _funcPerformance[item] / _funcs[item] + " milliseconds");
			}
			_sw.WriteLine();
		}
		_sw.WriteLine();
		_funcs.Clear();
	}

	public void LogFunction(string func)
	{
		if (!_funcs.ContainsKey(func))
		{
			_funcs.Add(func, 0);
		}
		_funcs[func]++;
	}

	public void LogFunction(string func, long milliseconds)
	{
		if (!_funcPerformance.ContainsKey(func))
		{
			_funcPerformance[func] = 0L;
		}
		_funcPerformance[func] += milliseconds;
	}

	public void Dispose()
	{
		_sw.Close();
		_sw.Dispose();
	}
}
