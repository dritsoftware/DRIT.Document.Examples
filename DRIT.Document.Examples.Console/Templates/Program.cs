using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var templatePath = ExampleSupport.InputPath("report-template.dotx");
if (!System.IO.File.Exists(templatePath))
{
	throw new InvalidOperationException(
		$"Add a DOCX template at '{templatePath}' before running this example.");
}

// Load the template as the starting document for a generated report.
var document = new Document(templatePath);
document.AttachedTemplate = templatePath;
document.Sections[0].Body.AddParagraph("Generated from the report template.");

// Save the completed report separately from the source template.
document.Save(ExampleSupport.OutputPath("Templates.docx"));
