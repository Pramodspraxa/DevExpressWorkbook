using DevExpress.Spreadsheet;
using DevExpressWorkbookApi;
using NLog;
using NLog.Extensions.Logging;

var builder = WebApplication.CreateBuilder(args);

builder.Logging.ClearProviders();
builder.Logging.AddNLog();

LogManager.LoadConfiguration("NLog.config");

var app = builder.Build();
var logger = LogManager.GetCurrentClassLogger();

// Endpoint to create a new workbook
app.MapGet("/generate-workbook/{templateType?}", async (HttpContext context, string? templateType) =>
{
	try
	{
		var columns = new string[30] { "Date", "HSCode", "ProductDescription", "Importer", "Exporter", "RelatedParty", "StdQty", "StdUnit", "GrossWeight", "Quantity", "UnitRateUSD", "QuantityUnit", "Value", "OriginCountry", "OriginPort", "DestinationCountry", "DestinationPort", "BillLadingNo", "Mode", "Measurment", "Tax", "DeliveryPortNameNew", "TEU", "FreightTermNew", "MarksNumber", "ImporterAdd1", "ExporterAdd1", "RelatedPartyAdd1", "HS4HS8Description", "CountryName" };
		int ChunkSize = 5000;
		int totalRecords = 60000;
		string fileName = "SampleData25.xlsx";

		templateType = string.IsNullOrWhiteSpace(templateType) ? "new" : "template-based";
		logger.Trace($"Starting workbook generation for TemplateType = {templateType}");

		var rawData = new List<GlobalDownloadMappingData>();

		for (int i = 1; i <= totalRecords; i++)
		{
			rawData.Add(new GlobalDownloadMappingData()
			{
				Date = new DateTime(2025, 01, 21),
				HSCode = $"HSCode {i}",
				ProductDescription = $"ProductDescription for record {i}",
				Importer = $"Test Importer for record {i}",
				Exporter = $"Test Exporter for record {i}",
				RelatedParty = $"Test RelatedParty for record {i}",
				StdQty = 157,
				StdUnit = $"UNT {i}",
				GrossWeight = 100 + i,
				Quantity = 500 + i,
				UnitRateUSD = 2 + i,
				QuantityUnit = "Package",
				Value = 3504 + i,
				OriginCountry = "China",
				DestinationCountry = "Test Destination",
				DestinationPort = "N/A",
				BillLadingNo = "N/A",
				Mode = "Air",
				Measurment = "N/A",
				Tax = "-",
				DeliveryPortNameNew = "-",
				TEU = "-",
				FreightTermNew = "-",
				MarksNumber = "-",
				ImporterAdd1 = $"Test Address {i}",
				ExporterAdd1 = $"Test Exp Address {i}",
				RelatedPartyAdd1 = $"Related Party Add - {i}",
				HS4HS8Description = $"HS 4 or HS 8 Desc {i}",
				CountryName = $"Main Country Source {i}"
			});
		}

		logger.Trace($"Dummy data generated. Populating workbook for TemplateType = {templateType}");
		bool applyFormatting = true;
		// Create a new workbook
		using (Workbook workbook = new Workbook())
		{
			Worksheet worksheet;
			if (templateType.Equals("template-based"))
			{
				logger.Trace($"workbook.LoadDocument Started for TemplateType = {templateType}.");
				Console.WriteLine("Worksheet load started");
				workbook.LoadDocument(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RawData.xlsx"));
				Console.WriteLine($"Workbook loaded");
				logger.Trace($"workbook.LoadDocument Completed for TemplateType = {templateType}.");
				workbook.BeginUpdate();
				worksheet = workbook.Worksheets[2];
			}
			else
			{
				worksheet = workbook.Worksheets[0];
			}
			int index = 0;
			Console.WriteLine($"Worksheet count -> {workbook.Worksheets.Count}");
			if (applyFormatting)
			{
				foreach (var item in columns)
				{
					worksheet[1, index].SetValue(item);
					worksheet[1, index].Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml("#16365c");
					worksheet[1, index].Font.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
					worksheet[1, index].Font.Bold = true;
					index++;
				}
				Style customStyle = workbook.Styles.Add("CustomStyle");
				customStyle.Font.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
				customStyle.Font.Bold = true;
				customStyle.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml("#16365c");

				CellRange range = worksheet.Range["A2:AD2"];
				range.Style = customStyle;
				worksheet.Columns[index - 1].WidthInPixels = 100;
				worksheet.Cells.Alignment.WrapText = true;
			}
			int rowIndex = 2; // Start from row 2 since row 1 is header
			int rawDataCount = rawData.Count;
			logger.Trace($"worksheet.Import Started for TemplateType = {templateType}.");
			for (int i = 0; i < rawDataCount; i += ChunkSize)
			{
				var chunk = rawData.Skip(i).Take(ChunkSize);
				using (var reader = FastMember.ObjectReader.Create(chunk, columns))
				{
					worksheet.Import(reader, false, rowIndex, 0);
				}
				rowIndex += chunk.Count();
			}
			if (applyFormatting)
			{
				worksheet.Columns["A"].NumberFormat = "dd-MMM-yyyy";
				DevExpress.Spreadsheet.CellRange rangeFilter = worksheet.Range["A2:AD2"];
				worksheet.AutoFilter.Apply(rangeFilter);
			}
			logger.Trace($"worksheet.Import Completed for TemplateType = {templateType}.");
			rawData = null;
			workbook.Calculate();
			workbook.EndUpdate();
			logger.Trace($"workbook.SaveDocument Started for TemplateType = {templateType}.");
			var data = await workbook.SaveDocumentAsync(DocumentFormat.Xlsx);
			workbook.Dispose();
			logger.Trace("Data saved to stream/ byte array");

			context.Response.Clear();
			context.Response.Headers["Content-Disposition"] = $"attachment; filename={fileName}";
			context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

			logger.Trace($"Writing workbook {fileName} to response for TemplateType = {templateType}.");
			await context.Response.Body.WriteAsync(data, 0, data.Length);
		}
	}
	catch (Exception ex)
	{
		logger.Error(ex, "Error occurred while generating the workbook.");
		context.Response.StatusCode = StatusCodes.Status500InternalServerError;
		// For debugging only — send full exception to client
    	await context.Response.WriteAsync($"Error occurred while generating the workbook: {ex.Message}\n{ex.StackTrace}");
		// await context.Response.WriteAsync("An error occurred while processing your request.");
	}
});

app.Run();