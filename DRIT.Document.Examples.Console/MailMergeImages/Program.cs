using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// An image field receives binary image data from the merge record.
document.Sections[0].Body.AddParagraph("Product image: {{Photo}}");

// Configure image sizing while the merge engine expands the field.
var result = document.MailMerge.ExecuteWithResult(
	new { Photo = ExampleSupport.OnePixelPng() },
	new DocumentMailMergeOptions
	{
		ImageOptions = new MailMergeImageOptions
		{
			WidthCm = 2,
			LockAspectRatio = true
		}
	});

Console.WriteLine($"Image fields processed: {result.FieldCount}");
document.Save(ExampleSupport.OutputPath("MailMergeImages.docx"));
