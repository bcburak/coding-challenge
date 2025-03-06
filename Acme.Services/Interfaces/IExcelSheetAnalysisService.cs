using Acme.Entities.Documents;
using Aspose.Cells;

namespace Acme.Services.Interfaces
{
    public interface IExcelSheetAnalysisService
    {
        Task SetupPromptsAsync();
        Task ChunkSheetsAsync(Document document, Workbook workbook, int chunkSize = -1);
    }
}
