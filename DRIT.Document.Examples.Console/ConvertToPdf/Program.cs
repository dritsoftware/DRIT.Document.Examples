using DRIT.Document;
using DRIT.Document.Examples.Shared;
using DRIT.Document.Export.Pdf;

var document = new Document();
document.AddSection().AddParagraph("Annual research report");
document.Sections[0].Body.AddParagraph("Exported directly to PDF with document properties and bookmarks.");

// Configure PDF-specific options at export time, then write the PDF output.
document.SaveAsPdf(
	ExampleSupport.OutputPath("ConvertToPdf.pdf"),
	new DocumentPdfSaveOptions { IncludeBookmarks = true });
