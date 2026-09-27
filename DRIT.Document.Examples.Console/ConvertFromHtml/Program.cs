using System.IO;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var inputPath = ExampleSupport.InputPath("sample.html");

// Create a small HTML fixture so the example is runnable without another application.
File.WriteAllText(inputPath,
	"<html><body><h1>Inventory</h1><p>Imported without Microsoft Word.</p>" +
	"<table><tr><td>SKU</td><td>Stock</td></tr><tr><td>NE-SRV-001</td><td>42</td></tr></table>" +
	"</body></html>");

var document = new Document(inputPath);

// Loading from the HTML path selects the HTML reader; saving as DOCX selects the Word writer.
document.Save(ExampleSupport.OutputPath("ConvertFromHtml.docx"));
