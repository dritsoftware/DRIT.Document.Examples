using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// Add content before and after an explicit page break.
document.AddSection().AddParagraph("Before the page break");
document.Sections[0].Body.AddParagraph().AppendBreak(BreakType.PageBreak);
document.Sections[0].Body.AddParagraph("After the page break");

// The DOCX extension selects the Word document writer.
document.Save(ExampleSupport.OutputPath("Breaks.docx"));
