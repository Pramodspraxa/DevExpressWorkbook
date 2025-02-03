namespace DevExpressWorkbookApi
{
	public class GlobalDownloadMappingData : GlobalDownloadMappingCommon
	{
		public DateTime Date { get; set; }
		public decimal StdQty { get; set; }
		public string StdUnit { get; set; }
		public string UnitRateINRStd { get; set; }
		public string UnitPrice { get; set; }
		public string Tax { get; set; }
		public string Mode { get; set; }
		public string Measurment { get; set; }
		public string BillLadingNo { get; set; }
		public string DeliveryPortNameNew { get; set; }
		public string TEU { get; set; }
		public string FreightTermNew { get; set; }
		public string MarksNumber { get; set; }
		public decimal UnitRateUSD { get; set; }
	}
	public class GlobalDownloadMappingCommon
	{
		public string CountryName { get; set; }
		public string HSCode { get; set; }
		public string ProductDescription { get; set; }
		public string Importer { get; set; }
		public string ImporterAdd1 { get; set; }
		public string ImporterAdd2 { get; set; }
		public string ImporterAdd3 { get; set; }
		public string ImporterAdd4 { get; set; }
		public string Exporter { get; set; }
		public string ExporterAdd1 { get; set; }
		public string ExporterAdd2 { get; set; }
		public string ExporterAdd3 { get; set; }
		public string ExporterAdd4 { get; set; }
		public string RelatedParty { get; set; }
		public string RelatedPartyAdd1 { get; set; }
		public string RelatedPartyAdd2 { get; set; }
		public string RelatedPartyAdd3 { get; set; }
		public string RelatedPartyAdd4 { get; set; }
		public decimal Value { get; set; }
		public string ValueFC { get; set; }
		public decimal Quantity { get; set; }
		public string QuantityUnit { get; set; }
		public decimal GrossWeight { get; set; }
		public string OriginCountry { get; set; }
		public string OriginPort { get; set; }
		public string DestinationCountry { get; set; }
		public string DestinationPort { get; set; }
		public string HS4HS8Description { get; set; }
		public string RecordID { get; set; }
	}
}
