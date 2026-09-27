using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Original content");
using var appended = new Document();
appended.AddSection().AddParagraph("Appended content");

// Import the second document while preserving its source formatting.
document.AppendDocument(appended, ImportFormatMode.KeepSourceFormatting);
document.Save(ExampleSupport.OutputPath("Combining.docx"));
