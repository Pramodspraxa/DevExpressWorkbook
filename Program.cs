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
		var columns = new string[10] { "Id", "Name", "Description", "Description2", "Description3", "Description4", "Description5", "Description6", "Description7", "Description8" };
		int ChunkSize = 5000;
		int totalRecords = 60000;
		string fileName = "SampleData.xlsx";

		templateType = string.IsNullOrWhiteSpace(templateType) ? "new" : "template-based";
		logger.Trace($"Starting workbook generation for TemplateType = {templateType}");

		var rawData = new List<GlobalData>();

		for (int i = 1; i <= totalRecords; i++)
		{
			rawData.Add(new GlobalData()
			{
				Id = i,
				Name = $"Name {i}",
				Description = $"Description for record {i}",
				Description2 = $"Description2 for record {i}",
				Description3 = $"Description3 for record {i}",
				Description4 = $"Description4 for record {i}",
				Description5 = $"Description5 for record {i}",
				Description6 = $"Description6 for record {i}",
				Description7 = $"Description7 for record {i}",
				Description8 = $"Description8 for record {i}",

			});
		}

		logger.Trace($"Dummy data generated. Populating workbook for TemplateType = {templateType}");

		// Create a new workbook
		using (Workbook workbook = new Workbook())
		{
			Worksheet worksheet;
			// Create an instance of XlsxLoadOptions
			//var options = new XlsxLoadOptions
			//{
			//	UseBufferedStreaming = true, // Enable buffered streaming
			//	BufferSize = 8192           // Set the buffer size (in bytes) for optimized loading
			//};
			if (templateType.Equals("template-based"))
			{
				logger.Trace($"worksheet.LoadDocument Started for TemplateType = {templateType}.");
				workbook.LoadDocument(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RawData.xlsx"));
				logger.Trace($"worksheet.LoadDocument Completed for TemplateType = {templateType}.");
				workbook.BeginUpdate();
				worksheet = workbook.Worksheets[2];
			}
			else
			{
				worksheet = workbook.Worksheets[0];
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
			logger.Trace($"worksheet.Import Completed for TemplateType = {templateType}.");
			rawData = null;
			workbook.Calculate();
			workbook.EndUpdate();
			using (var ms = new MemoryStream())
			{
				workbook.SaveDocument(ms, DocumentFormat.Xlsx);
				workbook.Dispose();
				ms.Seek(0, SeekOrigin.Begin);

				context.Response.Clear();
				context.Response.Headers["Content-Disposition"] = $"attachment; filename={fileName}";
				context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

				logger.Trace($"Writing workbook {fileName} to response for TemplateType = {templateType}.");
				await ms.CopyToAsync(context.Response.Body);
			}
		}
	}
	catch (Exception ex)
	{
		logger.Error(ex, "Error occurred while generating the workbook.");
		context.Response.StatusCode = StatusCodes.Status500InternalServerError;
		await context.Response.WriteAsync("An error occurred while processing your request.");
	}
});

app.Run();