using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// Field markers define the values and the nested region that the merge will replace.
document.Sections[0].Body.AddParagraph(
	"Customer: {{CustomerName}}\n" +
	"Item: {{Item}}\n" +
	"{{#Orders}}{{Item}}={{Quantity}};{{/Orders}}");

var result = document.MailMerge.ExecuteWithResult(new
{
	CustomerName = "Aiko Tanaka",
	Item = "NebulaCore Server",
	Orders = new[]
	{
		new { Item = "NebulaCore Server", Quantity = 1 },
		new { Item = "ArcLight Monitor", Quantity = 3 }
	}
});

// Execute the merge against an anonymous object and inspect the result counts.
Console.WriteLine($"Merged fields: {result.FieldCount}");
Console.WriteLine($"Expanded records: {result.RecordCount}");
document.Save(ExampleSupport.OutputPath("MailMerge.docx"));
