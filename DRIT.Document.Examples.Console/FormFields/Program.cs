using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("Name: ");

// Legacy form fields are added through the paragraph helper and expose typed data.
var formText = paragraph.AppendFormText("CustomerName", "Enter a name");
formText.MaximumLength = 80;
document.Bookmarks.Add(new Bookmark
{
	Name = "CustomerName",
	Paragraph = paragraph,
	Start = paragraph.GetText().Length,
	End = paragraph.GetText().Length
});

document.Save(ExampleSupport.OutputPath("FormFields.docx"));
