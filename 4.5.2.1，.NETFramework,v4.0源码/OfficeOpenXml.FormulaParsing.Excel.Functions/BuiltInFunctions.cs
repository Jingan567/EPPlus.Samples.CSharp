using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class BuiltInFunctions : FunctionsModule
{
	public BuiltInFunctions()
	{
		base.Functions["len"] = new Len();
		base.Functions["lower"] = new Lower();
		base.Functions["upper"] = new Upper();
		base.Functions["left"] = new Left();
		base.Functions["right"] = new Right();
		base.Functions["mid"] = new Mid();
		base.Functions["replace"] = new Replace();
		base.Functions["rept"] = new Rept();
		base.Functions["substitute"] = new Substitute();
		base.Functions["concatenate"] = new Concatenate();
		base.Functions["char"] = new CharFunction();
		base.Functions["exact"] = new Exact();
		base.Functions["find"] = new Find();
		base.Functions["fixed"] = new Fixed();
		base.Functions["proper"] = new Proper();
		base.Functions["search"] = new Search();
		base.Functions["text"] = new OfficeOpenXml.FormulaParsing.Excel.Functions.Text.Text();
		base.Functions["t"] = new T();
		base.Functions["hyperlink"] = new Hyperlink();
		base.Functions["value"] = new Value();
		base.Functions["int"] = new CInt();
		base.Functions["abs"] = new Abs();
		base.Functions["asin"] = new Asin();
		base.Functions["asinh"] = new Asinh();
		base.Functions["cos"] = new Cos();
		base.Functions["cosh"] = new Cosh();
		base.Functions["power"] = new Power();
		base.Functions["sign"] = new Sign();
		base.Functions["sqrt"] = new Sqrt();
		base.Functions["sqrtpi"] = new SqrtPi();
		base.Functions["pi"] = new Pi();
		base.Functions["product"] = new Product();
		base.Functions["ceiling"] = new Ceiling();
		base.Functions["count"] = new Count();
		base.Functions["counta"] = new CountA();
		base.Functions["countblank"] = new CountBlank();
		base.Functions["countif"] = new CountIf();
		base.Functions["countifs"] = new CountIfs();
		base.Functions["fact"] = new Fact();
		base.Functions["floor"] = new Floor();
		base.Functions["sin"] = new Sin();
		base.Functions["sinh"] = new Sinh();
		base.Functions["sum"] = new Sum();
		base.Functions["sumif"] = new SumIf();
		base.Functions["sumifs"] = new SumIfs();
		base.Functions["sumproduct"] = new SumProduct();
		base.Functions["sumsq"] = new Sumsq();
		base.Functions["stdev"] = new Stdev();
		base.Functions["stdevp"] = new StdevP();
		base.Functions["stdev.s"] = new Stdev();
		base.Functions["stdev.p"] = new StdevP();
		base.Functions["subtotal"] = new Subtotal();
		base.Functions["exp"] = new Exp();
		base.Functions["log"] = new Log();
		base.Functions["log10"] = new Log10();
		base.Functions["ln"] = new Ln();
		base.Functions["max"] = new Max();
		base.Functions["maxa"] = new Maxa();
		base.Functions["median"] = new Median();
		base.Functions["min"] = new Min();
		base.Functions["mina"] = new Mina();
		base.Functions["mod"] = new Mod();
		base.Functions["average"] = new Average();
		base.Functions["averagea"] = new AverageA();
		base.Functions["averageif"] = new AverageIf();
		base.Functions["averageifs"] = new AverageIfs();
		base.Functions["round"] = new Round();
		base.Functions["rounddown"] = new Rounddown();
		base.Functions["roundup"] = new Roundup();
		base.Functions["rand"] = new Rand();
		base.Functions["randbetween"] = new RandBetween();
		base.Functions["rank"] = new Rank();
		base.Functions["rank.eq"] = new Rank();
		base.Functions["rank.avg"] = new Rank(isAvg: true);
		base.Functions["quotient"] = new Quotient();
		base.Functions["trunc"] = new Trunc();
		base.Functions["tan"] = new Tan();
		base.Functions["tanh"] = new Tanh();
		base.Functions["atan"] = new Atan();
		base.Functions["atan2"] = new Atan2();
		base.Functions["atanh"] = new Atanh();
		base.Functions["acos"] = new Acos();
		base.Functions["acosh"] = new Acosh();
		base.Functions["var"] = new Var();
		base.Functions["varp"] = new VarP();
		base.Functions["large"] = new Large();
		base.Functions["small"] = new Small();
		base.Functions["degrees"] = new Degrees();
		base.Functions["isblank"] = new IsBlank();
		base.Functions["isnumber"] = new IsNumber();
		base.Functions["istext"] = new IsText();
		base.Functions["isnontext"] = new IsNonText();
		base.Functions["iserror"] = new IsError();
		base.Functions["iserr"] = new IsErr();
		base.Functions["error.type"] = new ErrorType();
		base.Functions["iseven"] = new IsEven();
		base.Functions["isodd"] = new IsOdd();
		base.Functions["islogical"] = new IsLogical();
		base.Functions["isna"] = new IsNa();
		base.Functions["na"] = new Na();
		base.Functions["n"] = new N();
		base.Functions["if"] = new If();
		base.Functions["iferror"] = new IfError();
		base.Functions["ifna"] = new IfNa();
		base.Functions["not"] = new Not();
		base.Functions["and"] = new And();
		base.Functions["or"] = new Or();
		base.Functions["true"] = new True();
		base.Functions["false"] = new False();
		base.Functions["address"] = new Address();
		base.Functions["hlookup"] = new HLookup();
		base.Functions["vlookup"] = new VLookup();
		base.Functions["lookup"] = new Lookup();
		base.Functions["match"] = new Match();
		base.Functions["row"] = new Row
		{
			SkipArgumentEvaluation = true
		};
		base.Functions["rows"] = new Rows
		{
			SkipArgumentEvaluation = true
		};
		base.Functions["column"] = new Column
		{
			SkipArgumentEvaluation = true
		};
		base.Functions["columns"] = new Columns
		{
			SkipArgumentEvaluation = true
		};
		base.Functions["choose"] = new Choose();
		base.Functions["index"] = new Index();
		base.Functions["indirect"] = new Indirect();
		base.Functions["offset"] = new Offset
		{
			SkipArgumentEvaluation = true
		};
		base.Functions["date"] = new Date();
		base.Functions["today"] = new Today();
		base.Functions["now"] = new Now();
		base.Functions["day"] = new Day();
		base.Functions["month"] = new Month();
		base.Functions["year"] = new Year();
		base.Functions["time"] = new Time();
		base.Functions["hour"] = new Hour();
		base.Functions["minute"] = new Minute();
		base.Functions["second"] = new Second();
		base.Functions["weeknum"] = new Weeknum();
		base.Functions["weekday"] = new Weekday();
		base.Functions["days360"] = new Days360();
		base.Functions["yearfrac"] = new Yearfrac();
		base.Functions["edate"] = new Edate();
		base.Functions["eomonth"] = new Eomonth();
		base.Functions["isoweeknum"] = new IsoWeekNum();
		base.Functions["workday"] = new Workday();
		base.Functions["networkdays"] = new Networkdays();
		base.Functions["networkdays.intl"] = new NetworkdaysIntl();
		base.Functions["datevalue"] = new DateValue();
		base.Functions["timevalue"] = new TimeValue();
		base.Functions["dget"] = new Dget();
		base.Functions["dcount"] = new Dcount();
		base.Functions["dcounta"] = new DcountA();
		base.Functions["dmax"] = new Dmax();
		base.Functions["dmin"] = new Dmin();
		base.Functions["dsum"] = new Dsum();
		base.Functions["daverage"] = new Daverage();
		base.Functions["dvar"] = new Dvar();
		base.Functions["dvarp"] = new Dvarp();
	}
}
