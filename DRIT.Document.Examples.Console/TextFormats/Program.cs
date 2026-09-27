using DRIT.Document;
using DRIT.Document.Examples.Shared;

var markdownPath = ExampleSupport.InputPath("sample.md");
if (!System.IO.File.Exists(markdownPath))
{
	System.IO.File.WriteAllText(markdownPath, "# Quarterly report\n\nThe report was imported from Markdown.");
}

// The load options select the Markdown reader while the output extension selects DOCX.
var document = new Document(markdownPath, new MarkdownLoadOptions());
document.Save(ExampleSupport.OutputPath("TextFormats.docx"));
