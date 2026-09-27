using DRIT.Document;
using DRIT.Document.Examples.Shared;
using DRIT.Document.Export.Pdf;

var document = new Document();
document.AddSection().AddParagraph("This PDF requires a password to open.");

// PDF encryption is configured on the PDF save options, not on the DOCX document.
document.SaveAsPdf(
	ExampleSupport.OutputPath("PdfSecurity.pdf"),
	new DocumentPdfSaveOptions
	{
			EncryptionDetails = new PdfEncryptionDetails("user-password", "owner-password")
	});
