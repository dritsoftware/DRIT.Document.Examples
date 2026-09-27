using System;
using System.IO;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var inputPath = ExampleSupport.InputPath("legacy.doc");
if (!File.Exists(inputPath))
{
	Console.WriteLine($"Add a binary DOC fixture at {inputPath} to run this import example.");
	return;
}

// The extension selects the legacy binary DOC reader.
using var document = new Document(inputPath);
document.Save(ExampleSupport.OutputPath("ConvertFromDoc.docx"));
