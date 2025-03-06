using Acme.Entities.Documents;

namespace Acme.Services.Interfaces
{
    public interface IDocumentAnalysisService
    {
        Task SetupPromptsAsync();
        Task<string?> GenerateTopicsFromOverviewsAsync(List<Page> pages);
        Task GenerateOverviewsAsync(Document document, bool forHfMonitoring = false);
        Task DetectSectionTitlesAsync(Document document);
        void GenerateTocUsingModelToc(Document document, string modelToc);
    }
}
