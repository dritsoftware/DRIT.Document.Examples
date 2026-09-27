using DRIT.Document.Examples.Shared;
using DRIT.Document;

var document = new Document();

// Each section has its own body and page setup.
document.AddSection().AddParagraph("First section");
document.AddSection().AddParagraph("Second section with independent content");

document.Save(ExampleSupport.OutputPath("Sections.docx"));
