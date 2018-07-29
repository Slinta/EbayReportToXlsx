using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EbayParser {


	public class RoughPayPalReport {
		public string Date;

		public string Time;
		public string TimeZone;
		public string Name;
		public string Type;
		public string Currency;
		public string Gross;
		public string ItemTitle;
		public string ShippingandHandlingAmount;
		
	}

	public class EbayTransaction {
		public DateTime date;
		public string sellerName;
		public string originalCurrency;
		public float usdValue;
		public string itemName;
		public float usdShippingCost;
		public Refund refund;
	}

	public class Refund {
		public Refund(string value, string datetime) {
			date = DateTime.ParseExact(datetime, "dd. MM. yyyyHH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
			usdValue = float.Parse(value.Replace(",", ""));
		}

		public float usdValue;
		public DateTime date;

	}
}
