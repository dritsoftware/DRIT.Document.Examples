using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.Sections[0].Body.AddParagraph("Order lines:\n{{#Items}}{{Sku}} = {{Quantity}};{{/Items}}");

// A collection assigned to Items expands the region once per record.
document.MailMerge.Execute(new
{
	Items = new[]
	{
		new { Sku = "NE-SRV-001", Quantity = 42 },
		new { Sku = "NE-MON-002", Quantity = 135 }
	}
});

document.Save(ExampleSupport.OutputPath("MailMergeRanges.docx"));
