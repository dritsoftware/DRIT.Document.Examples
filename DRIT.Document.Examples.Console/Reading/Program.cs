using DRIT.Document.Examples.Shared;
using System;
using DRIT.Document;

var path = ExampleSupport.OutputPath("Reading.docx");
var document = new Document();

// Save a source file first, then load it through the same public Document API.
document.AddSection().AddParagraph("Reading and text extraction");
document.Sections[0].Body.AddParagraph("The same document can be loaded again without Microsoft Word.");
document.Save(path);

// GetText exposes the loaded document as plain text.
using var loaded = new Document(path);
Console.WriteLine(loaded.GetText());
