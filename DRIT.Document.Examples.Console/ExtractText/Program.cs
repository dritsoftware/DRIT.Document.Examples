using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Extracted text");
document.Sections[0].Body.AddParagraph("GetText returns the document content as plain text.");

// Read the text representation before saving the document for inspection.
Console.WriteLine(document.GetText());
document.Save(ExampleSupport.OutputPath("ExtractText.docx"));
