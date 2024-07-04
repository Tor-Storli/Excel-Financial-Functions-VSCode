
// using System.Reflection;
// using Excel.FinancialFunctions;


// partial class Program
// { 
// class MACRSDeprciationTable
// {
//     public static Dictionary<Tuple<int, string>, double> table = new()
//     {
// 	{ Tuple.Create(1, "3-year"),0.3333},
// 	{ Tuple.Create(2, "3-year"),0.4445},
// 	{ Tuple.Create(3, "3-year"),0.1481},
// 	{ Tuple.Create(4, "3-year"),0.0741},
// 	{ Tuple.Create(1, "5-year"),0.2},
// 	{ Tuple.Create(2, "5-year"),0.32},
// 	{ Tuple.Create(3, "5-year"),0.192},
// 	{ Tuple.Create(4, "5-year"),0.1152},
// 	{ Tuple.Create(5, "5-year"),0.1152},
// 	{ Tuple.Create(6, "5-year"),0.0576},
// 	{ Tuple.Create(1, "7-year"),0.1429},
// 	{ Tuple.Create(2, "7-year"),0.2449},
// 	{ Tuple.Create(3, "7-year"),0.1749},
// 	{ Tuple.Create(4, "7-year"),0.1249},
// 	{ Tuple.Create(5, "7-year"),0.0893},
// 	{ Tuple.Create(6, "7-year"),0.0892},
// 	{ Tuple.Create(7, "7-year"),0.0893},
// 	{ Tuple.Create(8, "7-year"),0.0446},
// 	{ Tuple.Create(1, "10-year"),0.1},
// 	{ Tuple.Create(2, "10-year"),0.18},
// 	{ Tuple.Create(3, "10-year"),0.144},
// 	{ Tuple.Create(4, "10-year"),0.1152},
// 	{ Tuple.Create(5, "10-year"),0.0922},
// 	{ Tuple.Create(6, "10-year"),0.0737},
// 	{ Tuple.Create(7, "10-year"),0.0655},
// 	{ Tuple.Create(8, "10-year"),0.0655},
// 	{ Tuple.Create(9, "10-year"),0.0656},
// 	{ Tuple.Create(10, "10-year"),0.0655},
// 	{ Tuple.Create(11, "10-year"),0.0328},
// 	{ Tuple.Create(1, "15-year"),0.05},
// 	{ Tuple.Create(2, "15-year"),0.095},
// 	{ Tuple.Create(3, "15-year"),0.0855},
// 	{ Tuple.Create(4, "15-year"),0.077},
// 	{ Tuple.Create(5, "15-year"),0.0693},
// 	{ Tuple.Create(6, "15-year"),0.0623},
// 	{ Tuple.Create(7, "15-year"),0.059},
// 	{ Tuple.Create(8, "15-year"),0.059},
// 	{ Tuple.Create(9, "15-year"),0.0591},
// 	{ Tuple.Create(10, "15-year"),0.059},
// 	{ Tuple.Create(11, "15-year"),0.0591},
// 	{ Tuple.Create(12, "15-year"),0.059},
// 	{ Tuple.Create(13, "15-year"),0.0591},
// 	{ Tuple.Create(14, "15-year"),0.059},
// 	{ Tuple.Create(15, "15-year"),0.0591},
// 	{ Tuple.Create(16, "15-year"),0.0295},
// 	{ Tuple.Create(1, "20-year"),0.0375},
// 	{ Tuple.Create(2, "20-year"),0.07219},
// 	{ Tuple.Create(3, "20-year"),0.06677},
// 	{ Tuple.Create(4, "20-year"),0.06177},
// 	{ Tuple.Create(5, "20-year"),0.05713},
// 	{ Tuple.Create(6, "20-year"),0.05285},
// 	{ Tuple.Create(7, "20-year"),0.04888},
// 	{ Tuple.Create(8, "20-year"),0.04522},
// 	{ Tuple.Create(9, "20-year"),0.04462},
// 	{ Tuple.Create(10, "20-year"),0.04461},
// 	{ Tuple.Create(11, "20-year"),0.04462},
// 	{ Tuple.Create(12, "20-year"),0.04461},
// 	{ Tuple.Create(13, "20-year"),0.04462},
// 	{ Tuple.Create(14, "20-year"),0.04461},
// 	{ Tuple.Create(15, "20-year"),0.04462},
// 	{ Tuple.Create(16, "20-year"),0.04461},
// 	{ Tuple.Create(17, "20-year"),0.04462},
// 	{ Tuple.Create(18, "20-year"),0.04461},
// 	{ Tuple.Create(19, "20-year"),0.04462},
// 	{ Tuple.Create(20, "20-year"),0.04461},
// 	{ Tuple.Create(21, "20-year"),0.02231}
// 	};

//     public static double GetRate(int year, string type)
//     {
//         return table.TryGetValue(Tuple.Create(year, type), out double rate) ? rate : 0.0;
//     }

// 	public static int GetPctCount(string type){
// 		int countMacrsPCT = 0;

// 		foreach (var key in table.Keys)
//         {
//             if (key.Item2 == type){
//                 countMacrsPCT++;
// 			}	
//         }
// 		return countMacrsPCT;
// 	}
// }

// class tvmParameters
// {
// 	public required string excelfunction { get; set; }
// 	public double cost { get; set; }
// 	public double salvageValue { get; set; }
// 	public int life { get; set; }
// 	public int period { get; set; }
// 	public double factor { get; set; }	
// 	public double loanAmount { get; set; }
//     public double annualInterestRate { get; set; }
//     public double periodicInterestRate { get; set; }
//     public int numberOfPeriods { get; set; }
//     public int periodsPerYear { get; set; }
//     public int years { get; set; }
//     public double pmt_value { get; set; }
//     public double pv_value { get; set; }
//     public double fv_value { get; set; }
//     public int start_period { get; set; }
//     public int end_period { get; set; }
//     public int beg_end { get; set; }
// 	public required string macrs { get; set; }
// }

// static tvmParameters ParseArguments(string[] args)
// {
//         var p = new tvmParameters
//         {
//             // Initialize default values
//             excelfunction = "Pv", //  p.excelfunction = str; // default is Present Value
// 			cost = 0,
// 			salvageValue = 0, // p.salvageValue = double
//             life = 1, // p.life = int
// 			period = 1, // p.period = int
// 			factor = 2, // p.factor = double // default set as Double Declining Balance	Factor	
// 			loanAmount = 0,
//             annualInterestRate = 0,
//             periodicInterestRate = 0,
//             numberOfPeriods = 0,
//             periodsPerYear = 0,
//             years = 0,
//             pmt_value = 0,
//             pv_value = 0,
//             fv_value = 0,
//             start_period = 0,
//             end_period = 0,
//             beg_end = 0,
// 			macrs = "3-year"
//         };

//     // Parse command-line arguments
//     for (int i = 0; i < args.Length; i++)
//     {
//         switch (args[i])
//         {
// 		case "--excelfunction":
// 		if (i + 1 < args.Length)
// 			p.excelfunction = args[i + 1];
// 		break;

// 		case "--cost":
// 		if (i + 1 < args.Length)
// 			p.cost = double.Parse(args[i + 1]);
// 		break;

// 		case "--salvageValue":
// 		if (i + 1 < args.Length)
// 			p.salvageValue = double.Parse(args[i + 1]);
// 		break;

// 		case "--life":
// 		if (i + 1 < args.Length)
// 			p.life = int.Parse(args[i + 1]);
// 		break;

// 		case "--period":
// 		if (i + 1 < args.Length)
// 			p.period = int.Parse(args[i + 1]);
// 		break;	

// 		case "--factor":
// 		if (i + 1 < args.Length)
// 			p.factor = double.Parse(args[i + 1]);
// 		break;	

// 		case "--loanAmount":
// 		if (i + 1 < args.Length)
// 			p.loanAmount = double.Parse(args[i + 1]);
// 		break;

// 		case "--annualInterestRate":
// 		if (i + 1 < args.Length)
// 			p.annualInterestRate = double.Parse(args[i + 1]);
// 		break;

// 		case "--periodicInterestRate":
// 		if (i + 1 < args.Length)
// 			p.periodicInterestRate = double.Parse(args[i + 1]);
// 		break;

// 		case "--numberOfPeriods":
// 		if (i + 1 < args.Length)
// 			p.numberOfPeriods = int.Parse(args[i + 1]);
// 		break;

// 		case "--periodsPerYear":
// 		if (i + 1 < args.Length)
// 			p.periodsPerYear = int.Parse(args[i + 1]);
// 		break;

// 		case "--years":
// 		if (i + 1 < args.Length)
// 			p.years = int.Parse(args[i + 1]);
// 		break;

// 		case "--pmt_value":
// 		if (i + 1 < args.Length)
// 			p.pmt_value  = double.Parse(args[i + 1]);
// 		break;

// 		case "--pv_value":
// 		if (i + 1 < args.Length)
// 			p.pv_value   = double.Parse(args[i + 1]);
// 		break;

// 		case "--fv_value":
// 		if (i + 1 < args.Length)
// 			p.fv_value   = double.Parse(args[i + 1]);
// 		break;

// 		case "--start_period":
// 		if (i + 1 < args.Length)
// 			p.start_period = int.Parse(args[i + 1]);
// 		break;

// 		case "--end_period":
// 		if (i + 1 < args.Length)
// 			p.end_period = int.Parse(args[i + 1]);
// 		break;

// 		case "--beg_end":
// 		if (i + 1 < args.Length)
// 			p.beg_end = int.Parse(args[i + 1]);
// 		break;

// 		case "--macrsGroup":
// 		if (i + 1 < args.Length)
// 			p.macrs = args[i + 1];
// 		break;
//         }
//     }

//     return p;
// }

// static void Main(string[] args)
// {

//      var p = ParseArguments(args);

//      Console.WriteLine(p.excelfunction);

//      switch (p.excelfunction){
// 		case "Pv":
// 			double pv = Financial.Pv(p.periodicInterestRate, p.numberOfPeriods,p.pmt_value, p.fv_value, (PaymentDue)p.beg_end);
// 		    Console.WriteLine(pv);
// 		    break;
//         case "Fv":
// 			double fv = Financial.Fv(p.periodicInterestRate, p.numberOfPeriods, p.pmt_value, p.pv_value, (PaymentDue)p.beg_end);
// 		     Console.WriteLine(fv);
// 			break;
// 		case "Pmt":
// 			double pmt = Financial.Pmt(p.periodicInterestRate, p.numberOfPeriods, p.pv_value, p.fv_value, (PaymentDue)p.beg_end);
// 		    Console.WriteLine(pmt);
// 		    break;
// 		case "IPmt":
// 			double ipmt = Financial.IPmt(p.periodicInterestRate,p.start_period, p.numberOfPeriods, p.pv_value, p.fv_value, (PaymentDue)p.beg_end);
// 		    Console.WriteLine(ipmt);
// 		    break;
// 		case "PPmt":
// 			double ppmt = Financial.PPmt(p.periodicInterestRate,p.start_period, p.numberOfPeriods, p.pv_value, p.fv_value, (PaymentDue)p.beg_end);
// 		    Console.WriteLine(ppmt);
// 		    break;
// 		case "Rate":
// 			double rate = Financial.Rate(p.numberOfPeriods, p.pmt_value, p.pv_value, p.fv_value, (PaymentDue)p.beg_end);
// 		    Console.WriteLine(rate);
// 		    break;
// 		case "NPer":
// 			double nper = Financial.NPer(p.periodicInterestRate, p.pmt_value, p.pv_value, p.fv_value, (PaymentDue)p.beg_end);
// 		    Console.WriteLine(nper);
// 		    break;
// 		case "SLN":
// 			double sln = Financial.Sln(p.cost, p.salvageValue, p.life);
// 		    Console.WriteLine(sln);
// 		    break;	
// 		case "SYD":
// 			double syd = Financial.Syd(p.cost, p.salvageValue, p.life, p.period);
// 		    Console.WriteLine(syd);
// 		    break;		
// 		case "DDB":
// 			double ddb = Financial.Ddb(p.cost, p.salvageValue, p.life, p.period, p.factor);
// 		    Console.WriteLine(ddb);
// 		    break;
// 		case "MACRS":

// 		    string typeName = p.macrs;
// 			double accumDepr = 0;
// 			int macrsCount = MACRSDeprciationTable.GetPctCount(typeName);
// 			if (macrsCount > 0){

// 				Console.WriteLine($"{"Category",-10} {"Year",-5} {"Cost",-12} {"Percent",-10} {"Depreciation",-12} {"Accum-Depr",-14}");
// 				for (int n = 1; n <= macrsCount ; n++){
// 					double macrspct = MACRSDeprciationTable.GetRate(n, typeName);
// 					double deprAmount = p.cost * macrspct;
// 					accumDepr += deprAmount;
// 					 // Format the dollar amounts with thousand separators
//                     string formattedCost = p.cost.ToString("C0");
//                     string formattedDepreciation = deprAmount.ToString("C0");
// 					string formattedAccumDepr = accumDepr.ToString("C0");
// 					 Console.WriteLine($"{typeName,-10} {n,-5} {formattedCost,-12} {Math.Round(macrspct * 100, 2),-10} {formattedDepreciation,-12} {formattedAccumDepr, -14}");
// 					}
// 			}
// 			else{
// 				Console.WriteLine("Incomplete input data. Verify that you supplied the correct MACRS Type (i.e. 3-year | 5-year etc.)");
// 			}
// 			break;
// 	 }			
// 	}
//    }


