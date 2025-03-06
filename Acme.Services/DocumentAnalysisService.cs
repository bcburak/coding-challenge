using Acme.Entities.Documents;
using Acme.Services.Interfaces;

namespace Acme.Services;

public class DocumentAnalysisService : IDocumentAnalysisService
{
    public DocumentAnalysisService()
    {
    }

    public async Task SetupPromptsAsync()
    {
        await Task.CompletedTask;
    }

    public async Task<string?> GenerateTopicsFromOverviewsAsync(List<Page> pages)
    {
        await Task.CompletedTask;

        return "Topics";
    }

    public async Task GenerateOverviewsAsync(Document document, bool forHfMonitoring = false)
    {
        foreach (Page page in document.Pages)
        {
            page.Overview = "Overview";
        }

        await Task.CompletedTask;
    }

    public async Task DetectSectionTitlesAsync(Document document)
    {
        await Task.CompletedTask;
    }

    public void GenerateTocUsingModelToc(Document document, string modelToc)
    {
    }
}
