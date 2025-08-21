using DevExpress.Spreadsheet;
using DevExpressWorkbookApi;
using NLog;
using NLog.Extensions.Logging;
using System.Diagnostics;

var builder = WebApplication.CreateBuilder(args);

builder.Logging.ClearProviders();
builder.Logging.AddNLog();

try
{
	LogManager.LoadConfiguration("NLog.config");
	Console.WriteLine("NLog configuration loaded successfully");
}
catch (Exception ex)
{
	Console.WriteLine($"Failed to load NLog configuration: {ex.Message}");
	Console.WriteLine($"Exception: {ex}");
}

var app = builder.Build();
var logger = LogManager.GetCurrentClassLogger();

// Debug logging setup
Console.WriteLine($"Current Directory: {Directory.GetCurrentDirectory()}");
Console.WriteLine($"Base Directory: {AppDomain.CurrentDomain.BaseDirectory}");
Console.WriteLine($"NLog Configuration Loaded: {LogManager.Configuration != null}");

// Ensure logs directory exists
var logsPath = Path.Combine(Directory.GetCurrentDirectory(), "logs");
try
{
	if (!Directory.Exists(logsPath))
	{
		Directory.CreateDirectory(logsPath);
		Console.WriteLine($"Created logs directory: {logsPath}");
	}
	else
	{
		Console.WriteLine($"Logs directory exists: {logsPath}");
	}
	
	// Check permissions by creating a test file
	var testFile = Path.Combine(logsPath, "test.txt");
	File.WriteAllText(testFile, "test");
	File.Delete(testFile);
	Console.WriteLine("Directory permissions OK");
}
catch (Exception ex)
{
	Console.WriteLine($"Directory/permission issue: {ex.Message}");
}

// Test logging immediately
logger.Info("Application started - testing logging functionality");
logger.Trace("This is a trace message to test file logging");
Console.WriteLine("Logger test messages sent");

// Endpoint to create a new workbook
app.MapGet("/generate-workbook/{templateType?}", async (HttpContext context, string? templateType) =>
{
	try
	{
		var columns = new string[30] { "Date", "HSCode", "ProductDescription", "Importer", "Exporter", "RelatedParty", "StdQty", "StdUnit", "GrossWeight", "Quantity", "UnitRateUSD", "QuantityUnit", "Value", "OriginCountry", "OriginPort", "DestinationCountry", "DestinationPort", "BillLadingNo", "Mode", "Measurment", "Tax", "DeliveryPortNameNew", "TEU", "FreightTermNew", "MarksNumber", "ImporterAdd1", "ExporterAdd1", "RelatedPartyAdd1", "HS4HS8Description", "CountryName" };
		int ChunkSize = 5000;
		int totalRecords = 60000;
		string fileName = "SampleData25.xlsx";

		var totalStopwatch = Stopwatch.StartNew();
		templateType = string.IsNullOrWhiteSpace(templateType) ? "new" : "template-based";
		logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Starting workbook generation for TemplateType = {templateType}");

		var stepStopwatch = Stopwatch.StartNew();
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

		stepStopwatch.Stop();
		logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Dummy data generated in {stepStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}. Populating workbook for TemplateType = {templateType}");
		stepStopwatch.Restart();
		bool applyFormatting = true;
		// Create a new workbook
		using (Workbook workbook = new Workbook())
		{
			Worksheet worksheet;
			if (templateType.Equals("template-based"))
			{
				logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] workbook.LoadDocument Started for TemplateType = {templateType}.");
				Console.WriteLine("Worksheet load started");
				var loadStopwatch = Stopwatch.StartNew();
				workbook.LoadDocument(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RawData.xlsx"));
				loadStopwatch.Stop();
				Console.WriteLine($"Workbook loaded");
				logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] workbook.LoadDocument Completed in {loadStopwatch.Elapsed.ToString(@"mm\:ss\.fff")} for TemplateType = {templateType}.");
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
				logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Starting header formatting");
				var formatStopwatch = Stopwatch.StartNew();
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
				formatStopwatch.Stop();
				logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Header formatting completed in {formatStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");
			}
			int rowIndex = 2; // Start from row 2 since row 1 is header
			int rawDataCount = rawData.Count;
			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] worksheet.Import Started for TemplateType = {templateType}. Processing {rawDataCount} records in chunks of {ChunkSize}.");
			var importStopwatch = Stopwatch.StartNew();
			for (int i = 0; i < rawDataCount; i += ChunkSize)
			{
				var chunkStopwatch = Stopwatch.StartNew();
				var chunk = rawData.Skip(i).Take(ChunkSize);
				using (var reader = FastMember.ObjectReader.Create(chunk, columns))
				{
					worksheet.Import(reader, false, rowIndex, 0);
				}
				rowIndex += chunk.Count();
				chunkStopwatch.Stop();
				logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Processed chunk {(i / ChunkSize) + 1}/{(rawDataCount + ChunkSize - 1) / ChunkSize} ({chunk.Count()} records) in {chunkStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");
			}
			importStopwatch.Stop();
			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] All data import completed in {importStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");
			
			if (applyFormatting)
			{
				var postFormatStopwatch = Stopwatch.StartNew();
				worksheet.Columns["A"].NumberFormat = "dd-MMM-yyyy";
				DevExpress.Spreadsheet.CellRange rangeFilter = worksheet.Range["A2:AD2"];
				worksheet.AutoFilter.Apply(rangeFilter);
				postFormatStopwatch.Stop();
				logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Post-import formatting completed in {postFormatStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");
			}
			rawData = null;
			
			var calcStopwatch = Stopwatch.StartNew();
			workbook.Calculate();
			calcStopwatch.Stop();
			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Workbook calculation completed in {calcStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");
			
			workbook.EndUpdate();
			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] workbook.SaveDocument Started for TemplateType = {templateType}.");
			var saveStopwatch = Stopwatch.StartNew();
			var data = await workbook.SaveDocumentAsync(DocumentFormat.Xlsx);
			saveStopwatch.Stop();
			workbook.Dispose();
			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Data saved to stream/byte array in {saveStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");

			context.Response.Clear();
			context.Response.Headers["Content-Disposition"] = $"attachment; filename={fileName}";
			context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Writing workbook {fileName} to response for TemplateType = {templateType}.");
			var responseStopwatch = Stopwatch.StartNew();
			await context.Response.Body.WriteAsync(data, 0, data.Length);
			responseStopwatch.Stop();
			totalStopwatch.Stop();
			logger.Trace($"[{totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}] Response written in {responseStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}. Total execution time: {totalStopwatch.Elapsed.ToString(@"mm\:ss\.fff")}");
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