using Acme.Entities.Documents;
using Acme.Services.Interfaces;
using Aspose.Cells;

namespace Acme.Services;

public class ExcelSheetAnalysisService : IExcelSheetAnalysisService
{
    public ExcelSheetAnalysisService()
    {
    }

    public async Task SetupPromptsAsync()
    {
        await Task.CompletedTask;
    }

    public async Task ChunkSheetsAsync(
        Document document,
        Workbook workbook,
        int chunkSize = -1)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Page page = document[document.Pages.Count + 1];
            page.RawText = sheet.Name;
        }

        await Task.CompletedTask;
    }
}
