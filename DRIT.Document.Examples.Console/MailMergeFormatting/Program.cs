using System;
using System.Globalization;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// The template contains fields whose date, number, and Boolean output will be formatted.
document.Sections[0].Body.AddParagraph("{{When}} | {{Amount}} | {{Enabled}}");

// Formatting options are supplied with the merge operation, not embedded in the template.
var result = document.MailMerge.ExecuteWithResult(
	new
	{
		When = new DateTime(2026, 5, 12),
		Amount = 4899.0m,
		Enabled = true
	},
	new DocumentMailMergeOptions
	{
		Formatting = new MailMergeFormattingOptions
		{
			Culture = CultureInfo.GetCultureInfo("en-US"),
			DateTimeFormat = "yyyy-MM-dd",
			NumberFormat = "N2",
			BooleanTrueText = "YES",
			BooleanFalseText = "NO"
		}
	});

Console.WriteLine($"Formatted fields: {result.FieldCount}");
document.Save(ExampleSupport.OutputPath("MailMergeFormatting.docx"));
