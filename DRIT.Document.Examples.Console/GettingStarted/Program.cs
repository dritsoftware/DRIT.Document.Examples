using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

// Create a document, add a section and write two paragraphs.
var document = new Document();
var section = document.AddSection();
section.AddParagraph("NovaEdge Technologies - Employee Handbook 2026");
section.AddParagraph("This handbook sets out the company's policies and guidelines.");

var path = ExampleSupport.OutputPath("GettingStarted.docx");
document.Save(path);
Console.WriteLine($"Saved {path}");

// Loading the saved file is a useful minimal round-trip check.
using var loaded = new Document(path);
Console.WriteLine(loaded.GetText());
