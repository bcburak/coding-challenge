using Acme.Common;
using Acme.Common.Enums;
using Acme.Entities.Documents;
using Acme.Entities.Workflows;
using Acme.Entities.Workflows.Enums;
using Acme.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;
using File = Acme.Entities.Documents.File;

namespace Acme.Services.Tests;

public class DocumentProcessingServiceTests
{
    private readonly Mock<ILogger<DocumentProcessingService>> _mockLogger;
    private readonly Mock<IDocumentAnalysisService> _mockDocumentAnalysisService; // Could be useful if needed for feature developments
    private readonly Mock<IExcelSheetAnalysisService> _mockExcelSheetAnalysisService;
    // Should be create repositories and inject them here also
    private readonly DocumentProcessingService _service;

    private readonly BinaryContent _pdfSample;
    private readonly BinaryContent _csvSample;
    private readonly BinaryContent _excelSample;
    private readonly BinaryContent _imageSample;

    public DocumentProcessingServiceTests()
    {
        _mockLogger = new Mock<ILogger<DocumentProcessingService>>();
        _mockDocumentAnalysisService = new Mock<IDocumentAnalysisService>();
        _pdfSample = new(
            BinaryData.FromBytes(System.IO.File.ReadAllBytes("TestResources/MultiPageSearchable.pdf")),
            MediaType.ApplicationPdf);
        _csvSample = new(
            BinaryData.FromBytes(System.IO.File.ReadAllBytes("TestResources/csv.csv")),
            MediaType.TextCsv);
        _excelSample = new(
            BinaryData.FromBytes(System.IO.File.ReadAllBytes("TestResources/excel.xlsx")),
            MediaType.ApplicationVndMsExcel);
        _mockExcelSheetAnalysisService = new Mock<IExcelSheetAnalysisService>();
        _imageSample = new(
           BinaryData.FromBytes(System.IO.File.ReadAllBytes("TestResources/ManagerQs1.png")),
           MediaType.ImagePng);

        _service = new DocumentProcessingService(
            _mockLogger.Object,
            _mockDocumentAnalysisService.Object,
            _mockExcelSheetAnalysisService.Object);
    }

    [Fact]
    public async Task ProcessDocumentsAsync_ShouldLogWarning_WhenDocumentsListIsEmpty()
    {
        List<Document?> documents = new();

        await _service.ProcessDocumentsAsync(documents);

        _mockLogger.Verify(
            x => x.Log(
                LogLevel.Warning,
                It.IsAny<EventId>(),
                It.IsAny<It.IsAnyType>(),
                It.IsAny<System.Exception>(),
                It.IsAny<Func<It.IsAnyType, System.Exception?, string>>()
            ), Times.Once);
    }

    [Fact]
    public async Task ProcessDocumentsAsync_MultipleFiles_ShouldProcessCorrectly()
    {
        string? onDocProcessedIndicator = null;
        var documents = new List<Document?>
        {
            new() { MediaType = MediaType.ApplicationPdf, File = new File("test.pdf", _pdfSample.BinaryData.ToArray()) },
            new() { MediaType = MediaType.ApplicationVndMsExcel, File = new File("test.xlsx", _excelSample.BinaryData.ToArray()) },
            new() { MediaType = MediaType.ImagePng, File = new File("test.png", _imageSample.BinaryData.ToArray()) },
            new() { MediaType = MediaType.TextCsv, File = new File("test.csv", _csvSample.BinaryData.ToArray()) }
        };


        await _service.ProcessDocumentsAsync(
            documents,
            shouldGenerateOverviews: true,
            shouldDetectSectionTitles: true,
            onDocumentProcessed: async doc => await Task.Run(() => onDocProcessedIndicator += "success|"),
            cancellationToken: CancellationToken.None
        );

        Assert.Equal("success|success|success|success|", onDocProcessedIndicator);

        //Could be check other asseritons..
    }

    [Fact]
    public async Task LoadContextToWorkflowAsync_MultipleFiles_FillsWorkflow()
    {
        var workflow = CreateWorkflowWithBlocks();
        string? onDocProcessedIndicator = null;

        var documentInfos = new List<DocumentInfo>
        {
            new DocumentInfo { Document = new Document { Name = "doc1.pdf", MediaType = MediaType.ApplicationPdf, File = new File("MultiPageSearchable.pdf", _pdfSample.BinaryData.ToArray()) } },
            new DocumentInfo { Document = new Document { Name = "doc2.xlsx", MediaType = MediaType.ApplicationVndMsExcel, File = new File("excel.xlsx", _excelSample.BinaryData.ToArray()) } },
            new DocumentInfo { Document = new Document { Name = "doc3.png", MediaType = MediaType.ImagePng, File = new File("ManagerQs1.png", _imageSample.BinaryData.ToArray()) } },
            new DocumentInfo { Document = new Document { Name = "doc4.csv", MediaType = MediaType.TextCsv, File = new File("csv.csv", _csvSample.BinaryData.ToArray()) } }
        };

        await _service.LoadContextToWorkflowAsync(documentInfos, workflow, onDocumentProcessed: async document => await Task.Run(() => onDocProcessedIndicator += "success|"));


        //Could be check other asseritons..
        Assert.Null(onDocProcessedIndicator);
        Assert.NotEmpty(workflow.Blocks);
        Assert.Contains(workflow.Blocks, b => b.Documents.Any(d => d.Name == "doc1.pdf"));
    }

    private Workflow CreateWorkflowWithBlocks()
    {
        var workflow = new Workflow();

        var blocks = new List<Block>
    {
        new Block(workflow) { Name = "Block 1", ReplacementTag = "block1", ShouldUsePageImages = true, SupportedDocumentClasses = { "MultiPageSearchablePdf" } },
        new Block(workflow) { Name = "Block 2", ReplacementTag = "block2", ShouldUsePageImages = true },
        new Block(workflow) { Name = "Block 3", ReplacementTag = "block3", ShouldUsePageImages = true, SupportedDocumentClasses = { "MultiPageSearchablePdf", "xlsx" } },
        new Block(workflow) { Name = "Block 4", ReplacementTag = "block4", Type = BlockTypes.Merge }
    };

        foreach (var block in blocks)
        {
            _ = new RagSettings(block) { Type = RagTypes.WholeDocument };
        }

        return workflow;
    }

}
