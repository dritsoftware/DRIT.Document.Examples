using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Contents");
var introduction = document.Sections[0].Body.AddParagraph("Introduction");
introduction.ParagraphFormat.StyleName = "Heading1";
var details = document.Sections[0].Body.AddParagraph("Implementation details");
details.ParagraphFormat.StyleName = "Heading2";

// UpdateTableOfContents scans heading styles and populates the TOC entries.
document.UpdateTableOfContents();
document.Save(ExampleSupport.OutputPath("TableOfContents.docx"));
