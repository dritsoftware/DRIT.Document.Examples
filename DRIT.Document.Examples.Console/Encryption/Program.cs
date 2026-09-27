using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Password-protected document");

// Encryption belongs to DOCX save options and is applied when the package is written.
document.Save(
	ExampleSupport.OutputPath("Encryption.docx"),
	new DocumentSaveOptions
	{
		Encryption = new DocxEncryptionOptions
		{
			Password = "example-password",
			Mode = DocxEncryptionMode.Ecma376Agile
		}
	});
