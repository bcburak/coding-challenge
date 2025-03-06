using Acme.Common.Enums;
using Acme.Core;
using Acme.Entities.Documents;
using Acme.Entities.Workflows;
using Acme.Entities.Workflows.Enums;
using Acme.Services.Interfaces;
using Aspose.Cells;
using CsvHelper;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Text;
using PdfDocument = Aspose.Pdf.Document;

namespace Acme.Services;

public class DocumentProcessingService
{
    private const string ERROR_VECTOR_MODEL = "ERROR";

    private readonly ILogger<DocumentProcessingService> _logger;
    private readonly IDocumentAnalysisService _documentAnalysisService;
    private readonly IExcelSheetAnalysisService _excelSheetAnalysisService;

    // Made these services injected to the constructor to make the class more readeble and testable 
    //TODO: Could be improved by using a repository pattern to handle entity classes

    public DocumentProcessingService(
          ILogger<DocumentProcessingService> logger,
          IDocumentAnalysisService documentAnalysisService,
          IExcelSheetAnalysisService excelSheetAnalysisService)
    {
        _logger = logger;
        _documentAnalysisService = documentAnalysisService;
        _excelSheetAnalysisService = excelSheetAnalysisService;
    }

    //public async Task ProcessDocumentsAsync(List<Document?> documents, bool shouldGenerateOverviews = true, bool shouldDetectSectionTitles = true, Func<string?, Task>? onDocumentProcessed = null, CancellationToken cancellationToken = default)
    //{
    //    _logger.LogInformation("Processing {CountDocuments} documents.", documents.Count);
    //    await DocumentAnalysisService.SetupPromptsAsync().ConfigureAwait(false);
    //    await ExcelSheetAnalysisService.SetupPromptsAsync().ConfigureAwait(false);
    //    foreach (Document? document in documents)
    //    {
    //        if (document is null)
    //        {
    //            _logger.LogError("Document is null.");
    //            continue;
    //        }

    //        _logger.LogInformation("Processing document {DocumentName}.", document.Name);
    //        await ProcessPagesAsync(document, shouldGenerateOverviews, shouldDetectSectionTitles).ConfigureAwait(false);
    //        if (onDocumentProcessed is not null)
    //        {
    //            await onDocumentProcessed(document.Name).ConfigureAwait(false);
    //        }

    //        _logger.LogInformation("Document {DocumentName} processed.", document.Name);
    //    }

    //    _logger.LogInformation("Documents processed.");
    //}

    public async Task ProcessDocumentsAsync(List<Document?> documents, bool shouldGenerateOverviews = true, bool shouldDetectSectionTitles = true, Func<string?, Task>? onDocumentProcessed = null, CancellationToken cancellationToken = default)
    {
        if (documents is null || documents.Count == 0)
        {
            _logger.LogWarning("No documents provided for processing.");
            return;
        }

        _logger.LogInformation("Processing {CountDocuments} documents.", documents.Count);
        await SetupAnalysisServicesAsync();

        foreach (var document in documents)
        {
            _logger.LogInformation("Processing document {DocumentName}.", document!.Name);
            await ProcessPagesAsync(document, shouldGenerateOverviews, shouldDetectSectionTitles);

            if (onDocumentProcessed is not null)
                await onDocumentProcessed(document.Name);

            _logger.LogInformation("Document {DocumentName} processed.", document.Name);
        }

        _logger.LogInformation("Documents processed.");
    }

    private async Task SetupAnalysisServicesAsync()
    {
        await _documentAnalysisService.SetupPromptsAsync().ConfigureAwait(false);
        await _excelSheetAnalysisService.SetupPromptsAsync().ConfigureAwait(false);
    }
    public async Task LoadContextToWorkflowAsync(ICollection<DocumentInfo> documentInfos, Workflow workflow, int blockIndex = -1, Func<string?, Task>? onDocumentProcessed = null, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Loading context to workflow handler with {CountDocuments} documents.", documentInfos.Count);
        await _documentAnalysisService.SetupPromptsAsync().ConfigureAwait(false);
        await _excelSheetAnalysisService.SetupPromptsAsync().ConfigureAwait(false);
        List<Block> allBlocks = workflow.GetAllBlocks();
        foreach (Block block in allBlocks)
        {
            block.Documents.Clear();
        }

        foreach (DocumentInfo documentInfo in documentInfos)
        {
            if (documentInfo.Document is null)
            {
                _logger.LogError("Document {DocumentId} is null.", documentInfo.DocumentId);
                continue;
            }

            _logger.LogInformation("Loading document {DocumentName} to workflow.", documentInfo.Document.Name);
            AppendPagesToWorkflow(workflow, documentInfo.Document);
            AttachDocumentToBlocks(workflow, blockIndex, documentInfo);
            _logger.LogInformation("Document {DocumentName} processed.", documentInfo.Document.Name);
        }

        if (WorkflowUsesTopics(workflow))
        {
            _logger.LogInformation("Workflow uses topics, generating topics.");
            workflow.AllPagesTopics ??= await _documentAnalysisService.GenerateTopicsFromOverviewsAsync(workflow.AllPages).ConfigureAwait(false);
        }

        if (WorkflowUsesToc(workflow))
        {
            _logger.LogInformation("Workflow uses auto-detected TOC, generating TOC.");
            await GenerateTocsAsync(documentInfos, workflow, onDocumentProcessed);
        }
    }

    private static bool WorkflowUsesTopics(Workflow workflow)
    {
        return workflow.Blocks.Any(b => b.RagSettings.Any(r => r.Type is RagTypes.UseTopics)) || workflow.Children.Any(WorkflowUsesTopics);
    }

    private static bool WorkflowUsesToc(Workflow workflow)
    {
        return workflow.Blocks.Any(b => b.RagSettings.Any(r => r.Type is RagTypes.UseAutoDetectedTableOfContents)) || workflow.Children.Any(WorkflowUsesToc);
    }

    private async Task GenerateTocsAsync(ICollection<DocumentInfo> documentInfos, Workflow workflow, Func<string?, Task>? onDocumentProcessed)
    {
        var autoDetectedTocRags = workflow.Children.OrderBy(c => c.Order).SelectMany(c => c.Blocks.OrderBy(b => b.Order)).SelectMany(b => b.RagSettings).Where(r => r.Type is RagTypes.UseAutoDetectedTableOfContents).ToList();
        List<Document> documents = SelectDocumentsForToc(documentInfos, autoDetectedTocRags);
        if (documents.Count == 0)
        {
            _logger.LogInformation("No documents to generate TOC for.");
            return;
        }

        Dictionary<string, string?> tocModelsMap = GetTocModelsMap(documentInfos, autoDetectedTocRags);
        foreach (Document document in documents)
        {
            if (tocModelsMap.TryGetValue(document.Id, out string? tocModel) && tocModel is not null)
            {
                _logger.LogInformation("Generating TOC for document {DocumentName} with model {TocModel}.", document.Name, tocModel);
            }
            else
            {
                _logger.LogWarning("No TOC model found for document {DocumentName}.", document.Name);
                return;
            }

            _documentAnalysisService.GenerateTocUsingModelToc(document, tocModel);
            if (onDocumentProcessed is not null)
            {
                await onDocumentProcessed(document.Name).ConfigureAwait(false);
            }
        }
    }

    private Dictionary<string, string?> GetTocModelsMap(ICollection<DocumentInfo> documentInfos, List<RagSettings> autoDetectedTocRags)
    {
        _logger.LogInformation("Getting TOC models map for {CountDocuments} documents.", documentInfos.Count);
        Dictionary<string, string?> tocModelsMap = [];
        foreach (DocumentInfo documentInfo in documentInfos)
        {
            _logger.LogTrace("Searching for TOC models for document {DocumentName}.", documentInfo.Document?.Name);
            foreach (RagSettings rag in autoDetectedTocRags)
            {
                if (!documentInfo.CanBeProcessedForBlock(rag.Block ?? throw new InvalidOperationException("Block is null.")))
                {
                    continue;
                }

                if (string.IsNullOrEmpty(rag.TableOfContentsInput))
                {
                    _logger.LogWarning("Auto TOC RAG settings for block {BlockName} do not have a TOC model.", rag.Block?.Name);
                    continue;
                }

                tocModelsMap[documentInfo.DocumentId] = rag.TableOfContentsInput;
                break;
            }

            if (tocModelsMap.TryGetValue(documentInfo.DocumentId, out string? tocModel))
            {
                _logger.LogInformation("TOC model for document {DocumentName} found: {TocModel}.", documentInfo.Document?.Name, tocModel);
            }
            else
            {
                _logger.LogInformation("TOC model for document {DocumentName} not found.", documentInfo.Document?.Name);
            }
        }

        _logger.LogInformation("TOC models map for {CountDocuments} documents obtained.", documentInfos.Count);
        return tocModelsMap;
    }

    private List<Document> SelectDocumentsForToc(ICollection<DocumentInfo> documentInfos, List<RagSettings> autoDetectedTocRags)
    {
        List<Document> documents = [];
        foreach (DocumentInfo documentInfo in documentInfos)
        {
            if (documentInfo.Document is null)
            {
                _logger.LogError("Document {DocumentId} is null.", documentInfo.DocumentId);
                continue;
            }

            foreach (RagSettings rag in autoDetectedTocRags)
            {
                if (documentInfo.CanBeProcessedForBlock(rag.Block ?? throw new InvalidOperationException("Block is null.")))
                {
                    documents.Add(documentInfo.Document);
                    break;
                }
            }
        }

        return documents;
    }

    private async Task ProcessPagesAsync(Document document, bool shouldGenerateOverviews, bool shouldDetectSectionTitles)
    {
        switch (document.MediaType)
        {
            case MediaType.ApplicationPdf:
                await ProcessPdfPagesAsync(document, shouldGenerateOverviews, shouldDetectSectionTitles); //  i guess more readeable
                break;
            case MediaType.ImageJpeg:
            case MediaType.ImagePng:
                await ProcessMediaPagesAsync(document, false, shouldGenerateOverviews, shouldDetectSectionTitles);
                break;
            case MediaType.TextCsv:
                await ProcessCsvPagesAsync(document);
                break;
            case MediaType.ApplicationVndMsExcel:
            case MediaType.ApplicationVndOpenXmlFormatsOfficeDocumentSpreadsheetMlSheet:
                await ProcessExcelPagesAsync(document, shouldGenerateOverviews);
                break;
            default:
                throw new InvalidOperationException("Unknown document media type.");
        }
    }

    private void AppendPagesToWorkflow(Workflow workflow, Document document)
    {
        if (workflow.AllPages.Any(p => p.DocumentId == document.Id))
        {
            _logger.LogInformation("Document {DocumentName}'s pages are already in the workflow.", document.Name);
            return;
        }

        workflow.AllPages.AddRange(document.Pages);
        _logger.LogInformation("Document {DocumentName}'s pages added to the workflow.", document.Name);
    }

    private void AttachDocumentToBlocks(Workflow workflow, int blockIndex, DocumentInfo documentInfo)
    {
        foreach (Workflow child in workflow.Children)
        {
            AttachDocumentToBlocks(child, blockIndex, documentInfo);
        }

        if (blockIndex >= 0)
        {
            Block block = workflow[blockIndex];
            LoadDocumentToBlock(block, documentInfo);
            return;
        }

        foreach (Block block in workflow.Blocks)
        {
            LoadDocumentToBlock(block, documentInfo);
        }
    }

    private void LoadDocumentToBlock(Block block, DocumentInfo documentInfo)
    {
        if (documentInfo.Document is null)
        {
            _logger.LogError("Document {documentId} is null.", documentInfo.DocumentId);
            return;
        }

        if (!documentInfo.CanBeProcessedForBlock(block) || !block.ShouldUseDocuments())
        {
            _logger.LogDebug("Document {DocumentName} cannot be processed for block {BlockName}.", documentInfo.Document.Name, block.Name);
            return;
        }

        if (block.RagSettings.Count == 0)
        {
            _logger.LogWarning("Block {BlockName} does not have RAG settings.", block.Name);
            return;
        }

        if (block.Documents.Any(d => d.Id == documentInfo.DocumentId))
        {
            _logger.LogDebug("Document {DocumentName} is already attached to block {BlockName}.", documentInfo.Document.Name, block.Name);
            return;
        }

        _logger.LogInformation("Attaching document {DocumentName} to block {BlockName}.", documentInfo.Document.Name, block.Name);
        block.Documents.Add(documentInfo.Document);
    }

    private async Task ProcessCsvPagesAsync(Document document)
    {
        _logger.LogInformation("Processing CSV document {DocumentName}.", document.Name);
        if (document.File is null)
        {
            _logger.LogError("Document {DocumentName} does not have a File.", document.Name);
            return;
        }

        using var ms = new MemoryStream(document.File?.Bytes ?? throw new InvalidOperationException("Document file is null."));
        Page page = document[1];
        page.RawText = Encoding.UTF8.GetString(ms.ToArray());
        var stringBuilder = new StringBuilder(page.RawText.Length * 2);
        using var reader = new StreamReader(ms);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        string[] headerRecord = [];
        if (await csv.ReadAsync())
        {
            if (csv.HeaderRecord != null)
            {
                headerRecord = csv.HeaderRecord;
            }
            else
            {
                headerRecord = new string[csv.ColumnCount];
                for (int i = 0; i < csv.ColumnCount; i++)
                {
                    headerRecord[i] = csv.GetField(i) ?? string.Empty;
                }
            }
        }

        while (await csv.ReadAsync())
        {
            for (int i = 0; i < csv.ColumnCount; i++)
            {
                string? value = csv.GetField(i)?.Trim();
                if (!string.IsNullOrEmpty(value))
                {
                    stringBuilder.Append(headerRecord[i]);
                    stringBuilder.Append(": ");
                    stringBuilder.Append(csv.GetField(i));
                    stringBuilder.AppendLine();
                }
            }

            stringBuilder.AppendLine();
        }

        page.RawText = stringBuilder.ToString();
        double[]? vector = (await EmbeddingsConnector.GetResponseAsync([page.Overview ?? "An overview for this content is currently unavailable"]).ConfigureAwait(false)).FirstOrDefault();
        string modelName = page.Overview is null || vector is null ? ERROR_VECTOR_MODEL : EmbeddingsConnector.MODEL_NAME;
        page.EmbeddingVector = new(modelName, vector ?? []);
    }

    private async Task ProcessExcelPagesAsync(Document document, bool shouldGenerateOverviews)
    {
        _logger.LogInformation("Processing Excel document {DocumentName}.", document.Name);

        CheckIfFileIsNull(document);

        if (document.Pages.Count == 0)
        {
            using var ms = new MemoryStream(document.File?.Bytes ?? throw new InvalidOperationException("Document file is null."));
            var workbook = new Workbook(ms);
            await _excelSheetAnalysisService.ChunkSheetsAsync(document, workbook).ConfigureAwait(false);
        }
        else
        {
            _logger.LogInformation("Document {DocumentName} already has pages.", document.Name);
        }

        if (shouldGenerateOverviews)
        {
            await _documentAnalysisService.GenerateOverviewsAsync(document).ConfigureAwait(false);
        }

        await GetEmbeddingsAsync(document).ConfigureAwait(false);
    }

    private async Task ProcessMediaPagesAsync(Document document, bool shouldUseOcr, bool shouldGenerateOverviews, bool shouldDetectSectionTitles)
    {
        _logger.LogInformation("Processing media document {DocumentName}.", document.Name);
        CheckIfFileIsNull(document);

        Page page = document[1];
        page.Image ??= new(new BinaryData(document.File.Bytes), document.MediaType);
        if (!shouldUseOcr)
        {
            return;
        }

        _logger.LogInformation("Processing media document {DocumentName} with OCR.", document.Name);
        string? rawText = shouldUseOcr ? throw new NotImplementedException() : null;
        page.RawText ??= rawText?.Replace("\0", "");
        page.Text ??= rawText?.Replace("\0", "");
        if (shouldGenerateOverviews)
        {
            await _documentAnalysisService.GenerateOverviewsAsync(document).ConfigureAwait(false);
        }

        if (shouldDetectSectionTitles)
        {
            await _documentAnalysisService.DetectSectionTitlesAsync(document).ConfigureAwait(false);
        }

        if (IsPageEmbeddingValid(page))
        {
            _logger.LogDebug("Document {DocumentName} already has an embedding for page {PageNumber}.", document.Name, page.PageNumber);
            return;
        }

        if (string.IsNullOrWhiteSpace(page.Overview))
        {
            _logger.LogError("Failed to get overview for page {PageNumber}.", page.PageNumber);
        }

        double[]? vector = (await EmbeddingsConnector.GetResponseAsync([page.Overview ?? "An overview for this content is currently unavailable"]).ConfigureAwait(false)).FirstOrDefault();
        string modelName = page.Overview is null || vector is null ? ERROR_VECTOR_MODEL : EmbeddingsConnector.MODEL_NAME;
        page.EmbeddingVector = new(modelName, vector ?? []);
    }

    private async Task ProcessPdfPagesAsync(Document document, bool shouldGenerateOverviews, bool shouldDetectSectionTitles)
    {
        _logger.LogInformation("Processing PDF document {DocumentName}.", document.Name);
        CheckIfFileIsNull(document);

        await GetPdfPagesRawTextsAsync(document).ConfigureAwait(false);
        if (shouldGenerateOverviews)
        {
            await _documentAnalysisService.GenerateOverviewsAsync(document).ConfigureAwait(false);
        }

        if (shouldDetectSectionTitles)
        {
            await _documentAnalysisService.DetectSectionTitlesAsync(document).ConfigureAwait(false);
        }

        await GetEmbeddingsAsync(document).ConfigureAwait(false);
    }


    //will be refactor
    private async Task GetEmbeddingsAsync(Document document)
    {
        var pagesToEmbed = document.Pages.Where(p => !IsPageEmbeddingValid(p) && !string.IsNullOrWhiteSpace(p.Overview)).ToList();
        if (pagesToEmbed.Count == 0)
        {
            return;
        }
        _logger.LogInformation(
                      "Embedding {CountToEmbed} pages for document {DocumentName}.",
                      pagesToEmbed.Count, document.Name
                  );
        int countAlreadyEmbedded = document.Pages.Where(p => p.EmbeddingVector is not null).Count();
        int countErroneous = document.Pages.Where(p => p.EmbeddingVector?.Model == ERROR_VECTOR_MODEL).Count();
        _logger.LogInformation("Document {DocumentName} has {CountPages} pages, {CountAlreadyEmbedded} pages already embedded, {CountErroneous} pages with erroneous embeddings, {CountToEmbed} pages to embed.", document.Name, document.Pages.Count, countAlreadyEmbedded, countErroneous, pagesToEmbed.Count);


        var texts = pagesToEmbed.Select(p => p.Overview ?? $"An overview for this content is currently unavailable").ToList();
        List<double[]> allEmbeddings = [];
        int pagesInBatch = 50;
        for (int i = 0; i < texts.Count; i += pagesInBatch)
        {
            var batch = texts.Skip(i).Take(pagesInBatch).ToList();
            List<double[]> batchEmbeddings = await EmbeddingsConnector.GetResponseAsync(batch).ConfigureAwait(false);
            allEmbeddings.AddRange(batchEmbeddings);
        }

        for (int i = 0; i < pagesToEmbed.Count; i++)
        {
            double[]? vector = i < allEmbeddings.Count ? allEmbeddings[i] : null;
            string modelName = pagesToEmbed[i].Overview is null || vector is null ? ERROR_VECTOR_MODEL : EmbeddingsConnector.MODEL_NAME;
            if (modelName == ERROR_VECTOR_MODEL)
            {
                _logger.LogWarning("Failed to get embedding for page {PageNumber}.", pagesToEmbed[i].PageNumber);
            }

            EmbeddingVector embeddingVector = pagesToEmbed[i].EmbeddingVector ?? new(modelName, []);
            embeddingVector.Model = modelName;
            embeddingVector.Vector = vector ?? [];
            pagesToEmbed[i].EmbeddingVector = embeddingVector;
        }
    }

    private async Task GetPdfPagesRawTextsAsync(Document document)
    {
        CheckIfFileIsNull(document);
        using var ms = new MemoryStream(document.File.Bytes);
        using PdfDocument pdf = new(ms);
        Dictionary<int, string> pagesTexts = await PdfExtractor.ExtractPaginatedTextAsync(pdf, Enumerable.Range(1, pdf.Pages.Count)).ConfigureAwait(false);
        foreach ((int pageNumber, string pageText) in pagesTexts)
        {
            document[pageNumber].RawText = pageText.Replace("\0", "");
        }
    }

    private void CheckIfFileIsNull(Document document)
    {
        if (document.File is null)
        {
            _logger.LogError("Document {DocumentName} does not have a File.", document.Name);
            return;
        }
    }

    private static bool IsPageEmbeddingValid(Page page)
    {
        return page.EmbeddingVector is not null && page.EmbeddingVector.Model != ERROR_VECTOR_MODEL;
    }
}
