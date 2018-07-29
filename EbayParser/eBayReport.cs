using System;
using System.Collections;

namespace EbayParser {
	class eBayReport {
		//public string orderId;
		public Seller seller;
		public int itemIndex;
		public string purchaseDate;
		public int elapsedDays;
		public string price;
		public string quantity;
		public string specs;
		public string deliveryDate;
		public int etaDays;
		public string shipStatus;
		public bool feedBackNotLeft;
		public string thumbnail;
		public TrackingNo trackingNo;

	}
	class Seller {
		public string name;
		public string url;
	}
	class TrackingNo {
		public string name;
		public string url;
	}
}
