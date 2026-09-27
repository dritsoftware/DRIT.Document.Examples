using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var inputPath = ExampleSupport.InputPath("macro-enabled-source.docm");
if (!System.IO.File.Exists(inputPath))
{
	throw new InvalidOperationException(
		$"Add a macro-enabled source document at '{inputPath}' before running this example.");
}

// Loading a DOCM preserves the document's VBA project and embedded package parts.
var document = new Document(inputPath);
Console.WriteLine($"VBA modules: {document.VbaProject.Modules.Count}");

// Saving with the macro-enabled extension keeps the preserved references in the package.
document.Save(ExampleSupport.OutputPath("OleAndVbaReferences.docm"));
