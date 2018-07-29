using System;
using System.IO;
using OfficeOpenXml;
using FileHelpers;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EbayParser {
	class Program {
		public static List<RoughPayPalReport> payPalReports = new List<RoughPayPalReport>();
		public static Dictionary<string, int> PPRindexbyRename = new Dictionary<string, int>();
		public static List<EbayTransaction> purchases = new List<EbayTransaction>();
		public static bool debug = false;

		static void Main(string[] args) {
			Console.WriteLine("hello, this program will take a report from paypal in csv format,");
			Console.WriteLine("and then take all ebay transactions and insert them into a spreadsheet");
			Console.WriteLine("with costs translated into the currency of your paypal account, press enter");
			string launchOption = Console.ReadLine();
			
			if (launchOption == "debug") {
				debug = true;
			}
			Console.WriteLine("Please type in the name of the excel spreadsheet you wish to create");
			string spreadsheetName = Console.ReadLine();
			FileInfo excelFile = new FileInfo(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + spreadsheetName + ".xlsx");
			ExcelPackage excelPackage = new ExcelPackage(excelFile);
			string worksheetName = "eBay";
			if (excelPackage.Workbook.Worksheets[worksheetName] == null) {
				excelPackage.Workbook.Worksheets.Add(worksheetName);
			}
			else {
				Console.WriteLine("The file already exists and has a worksheet with the default name \"eBay\" in it.");
				Console.WriteLine("Specify a name of the new worksheet");
				worksheetName = Console.ReadLine();
				excelPackage.Workbook.Worksheets.Add(worksheetName);
			}

			try {
				excelPackage.Save();
			}
			catch {
				RedCW("The file is already open");
				Console.ReadLine();
				return;
			}

			DirectoryInfo currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
			FileInfo[] filesInCurrentDirectory = currentDirectory.GetFiles();
			List<List<eBayReport>> ebayReports = new List<List<eBayReport>>();
			foreach (FileInfo file in filesInCurrentDirectory) {
				if (file.Extension == ".json") {
					ebayReports.Add(JsonConvert.DeserializeObject<List<eBayReport>>(File.ReadAllText(file.FullName)));

				}
			}

			Console.WriteLine("Now specify the full name (with extension) of the CSV file");
			string csvFile = Console.ReadLine();
			string[] linesOfCSV;
			try {
				linesOfCSV = File.ReadAllLines(csvFile);
			}
			catch {
				RedCW("The file is already open");
				Console.ReadLine();
				return;
			}
			ExcelWorksheet sheet = excelPackage.Workbook.Worksheets[worksheetName];
			string[] initialLine = linesOfCSV[0].Split(',');
			FieldInfo[] fieldsInPaypalReport = typeof(RoughPayPalReport).GetFields();
			List<int> wantedPositions = new List<int>();
			for (int i = 0; i < initialLine.Length; i++) {
				foreach (FieldInfo field in fieldsInPaypalReport) {
					if (initialLine[i].Replace(" ", "").Replace("\"", "") == field.Name) {
						wantedPositions.Add(i);
						
						break;
					}
				}
			}

			FieldInfo[] fieldsInEBayReport = typeof(EbayTransaction).GetFields();
			FieldInfo[] fieldsInRefund = typeof(Refund).GetFields();

			for (int i = 0; i < fieldsInEBayReport.Length + fieldsInRefund.Length; i++) {
				if (i < fieldsInEBayReport.Length) {
					sheet.SetValue(1, i + 1, fieldsInEBayReport[i].Name);
				}
				else {
					sheet.SetValue(1, i + 1, fieldsInRefund[i - fieldsInEBayReport.Length].Name);
				}
			}

			

			for (int i = 1; i < linesOfCSV.Length; i++) {
				List<string> line = linesOfCSV[i].Split(new string[] { "\",\"" }, StringSplitOptions.None).ToList<string>();
				line[0] = line[0].Remove(0, 1);
				line[line.Count - 1] = line[line.Count - 1].Remove(line[line.Count - 1].Length - 1, 1);
				//List<string> adjustedLine = new List<string>();
				//foreach (string value in line) {
				//	if (value != "," && value != "") {
				//		adjustedLine.Add(value);
				//	}
				//}
				RoughPayPalReport report = new RoughPayPalReport();
				for (int indexOfSignificant = 0; indexOfSignificant < wantedPositions.Count; indexOfSignificant++) {

					fieldsInPaypalReport[indexOfSignificant].SetValue(report, line[wantedPositions[indexOfSignificant]].Replace("\"", ""));
					
					

				}
				payPalReports.Add(report);
			}
			Console.WriteLine("All parsing is complete, now you can rename the items or skip individual ones with enter");
			
			for (int i = 0; i < payPalReports.Count; i++) {
				switch (payPalReports[i].Type) {
					case "eBay Auction Payment": {
						EbayTransaction purchase = new EbayTransaction();
						purchase.date = DateTime.ParseExact(payPalReports[i].Date, "dd. MM. yyyy", System.Globalization.CultureInfo.InvariantCulture);
						purchase.time = DateTime.ParseExact(payPalReports[i].Time, "HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
						purchase.originalCurrency = payPalReports[i].Currency;
						if (purchase.originalCurrency != "USD") {
							purchase.usdValue = GetUsdValue(i);
						}
						else {
							purchase.usdValue = float.Parse(payPalReports[i].Gross.Replace(",", ""));
						}
						purchase.usdShippingCost = float.Parse(payPalReports[i].ShippingandHandlingAmount.Replace(",", ""));
						purchase.sellerName = payPalReports[i].Name;
						Console.WriteLine("Item: " + payPalReports[i].ItemTitle);
						Console.WriteLine("Leave empty to keep name or type new one");
						string s = Console.ReadLine();
						if (s.Trim() == "") {
							purchase.itemName = payPalReports[i].ItemTitle;
						}
						else {
							purchase.itemName = s.Trim();
						}
						PPRindexbyRename.Add(purchase.itemName, i);
						purchases.Add(purchase);
						break;
					}
					case "Payment Refund": {
						RoughPayPalReport currentReport = payPalReports[i];
						List<int> matchingPurchaseIndexesSeller = new List<int>();

						for (int j = 0; j < purchases.Count; j++) {
							if (purchases[j].sellerName == currentReport.Name) {
								matchingPurchaseIndexesSeller.Add(j);
							}
						}

						if (matchingPurchaseIndexesSeller.Count == 1) {

							purchases[matchingPurchaseIndexesSeller[0]].refund = RefundFound(purchases[matchingPurchaseIndexesSeller[0]],i);
						}
						else if (matchingPurchaseIndexesSeller.Count == 0) {
							Console.WriteLine("Found a refund with seller inconsistencies");
							Console.WriteLine("Item name: " + currentReport.ItemTitle);

							SellerInconsistencyHandling(i);


						}
						else {
							Console.WriteLine("Found more sales from one seller");
							Console.WriteLine("Item name: " + currentReport.ItemTitle);

							SellerInconsistencyHandling(i);
						}
						break;
					}
				}


			}
			
			for (int i = 0; i < purchases.Count; i++) {
				int line = 2 + i;
				for (int j = 0; j < fieldsInEBayReport.Length + fieldsInRefund.Length; j++) {
					if (j < fieldsInEBayReport.Length) {
						sheet.SetValue(line, j + 1, fieldsInEBayReport[j].GetValue(purchases[i]));
					}
					else if(purchases[i].refund != null) {
						sheet.SetValue(line, j + 1, fieldsInRefund[j - fieldsInEBayReport.Length].GetValue(purchases[i].refund));
					}
					
				}

				
				//sheet.SetValue(line, 1, purchases[i].date.ToShortDateString());
				//sheet.SetValue(line, 2, purchases[i].date.ToShortTimeString());


				//sheet.SetValue(line, 3, purchases[i].itemName);
				//sheet.SetValue(line, 4, purchases[i].originalCurrency);
				//sheet.SetValue(line, 5, purchases[i].sellerName);
				//sheet.SetValue(line, 6, purchases[i].usdShippingCost);
				//sheet.SetValue(line, 7, purchases[i].usdValue);
				//if (purchases[i].refund == null) {
				//	sheet.SetValue(line, 8, "NoRefunds");
				//}
				//else {
				//	sheet.SetValue(line, 9, purchases[i].refund.usdValue);
				//	sheet.SetValue(line, 10, purchases[i].refund.date);
				//}


			}
			excelPackage.Save();
		}

		public static float GetUsdValue(int pos) {
			string time = payPalReports[pos].Time;
			string date = payPalReports[pos].Date;
			if (payPalReports[pos - 2].Time == time && payPalReports[pos - 2].Date == date) {
				return float.Parse(payPalReports[pos - 2].Gross.Replace(",", ""));
			}
			if (payPalReports[pos + 1].Time == time && payPalReports[pos + 1].Date == date) {
				return float.Parse(payPalReports[pos + 1].Gross.Replace(",", ""));
			}
			throw new Exception("no conversion transaction found");
		}

		public static Refund RefundFound(EbayTransaction originalPurchase, int refundTransactionIndex ) {
			Console.WriteLine("Found a refund for item " + originalPurchase.itemName);
			if (debug) {
				Console.WriteLine("Original name was: " + payPalReports[PPRindexbyRename[originalPurchase.itemName]].Name);
			}

			string usdValue = "";
			if (originalPurchase.originalCurrency == "USD") {
				usdValue = payPalReports[refundTransactionIndex].Gross;
			}
			else {
				usdValue = GetUsdValue(refundTransactionIndex).ToString();
			}

			return new Refund(usdValue, payPalReports[refundTransactionIndex].Date + payPalReports[refundTransactionIndex].Time);
		}

		public static void SellerInconsistencyHandling(int refundIndex) {
			List<int> matchingPurchaseIndexesItem = new List<int>();
			for (int j = 0; j < purchases.Count; j++) {
				if (payPalReports[PPRindexbyRename[purchases[j].itemName]].Name == payPalReports[refundIndex].Name) {
					matchingPurchaseIndexesItem.Add(j);
				}
			}

			if (matchingPurchaseIndexesItem.Count == 1) {
				purchases[matchingPurchaseIndexesItem[0]].refund = RefundFound(purchases[matchingPurchaseIndexesItem[0]], refundIndex);
			}
			else if (matchingPurchaseIndexesItem.Count == 0) {
				Console.WriteLine("Nothing found");
				Console.WriteLine("No idea what to do now, skipping");
			}
			else {
				Dictionary<int, int> pprItemIndexToPurchaseIndex = new Dictionary<int, int>();
				List<int> pprItemIndexes = new List<int>();

				Console.WriteLine("Found more items with same name");
				Console.WriteLine("Indexes are: ");
				for (int indexindex = 0; indexindex < matchingPurchaseIndexesItem.Count; indexindex++) {
					int pprItemIndex = PPRindexbyRename[purchases[matchingPurchaseIndexesItem[indexindex]].itemName];
					pprItemIndexToPurchaseIndex.Add(pprItemIndex, matchingPurchaseIndexesItem[indexindex]);
					Console.Write(indexindex + ". " + pprItemIndex);
					pprItemIndexes.Add(pprItemIndex);
				}
				Console.WriteLine("Index of refund: " + refundIndex);
				pprItemIndexes.Sort();


				List<int> candidates = new List<int>();

				for (int k = 0; k < pprItemIndexes.Count; k++) {
					if (refundIndex > pprItemIndexes[k]) {
						candidates.Add(pprItemIndexes[k]);

					}

				}
				int savedFinalPos = candidates.Max();


				purchases[matchingPurchaseIndexesItem[0]].refund = RefundFound(purchases[pprItemIndexToPurchaseIndex[savedFinalPos]], refundIndex);




				Console.WriteLine("Matched {0} with {1} ", savedFinalPos, refundIndex);
			}

		}

		public static void RedCW(string text) {
			ConsoleColor c = Console.ForegroundColor;
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine(text);
			Console.ForegroundColor = c;
		}
	}
}
