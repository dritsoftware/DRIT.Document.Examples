using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.Sections[0].Body.AddParagraph("{{Name}} / {{Name}} / {{#Orders}}{{Item}}{{/Orders}}");

// Analyze does not mutate the template, so it can run before a batch merge.
var analysis = document.MailMerge.Analyze(new
{
	Name = "Ada",
	Orders = new[] { new { Item = "Book" } }
});

Console.WriteLine($"Template valid: {analysis.IsValid}");
Console.WriteLine($"Fields: {analysis.Fields.Count}; diagnostics: {analysis.Diagnostics.Count}");
document.Save(ExampleSupport.OutputPath("MailMergeAnalysis.docx"));
