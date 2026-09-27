using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("Customer name");

// Add a structured document tag to make this paragraph an editable named field.
paragraph.ContentControls.Add(new ContentControl
{
	Alias = "CustomerName",
	Title = "Customer name",
	Tag = "customer-name",
	Id = 1001,
	Type = ContentControlType.RichText
});

document.Save(ExampleSupport.OutputPath("ContentControls.docx"));
