using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var section = document.AddSection();
var paragraph = section.AddParagraph("Read the project details online.");

// A bookmark names a content range so fields and hyperlinks can target it.
document.Bookmarks.Add(new Bookmark
{
	Name = "ProjectDetails",
	Paragraph = paragraph,
	Start = 0,
	End = paragraph.GetText().Length
});

// External hyperlinks carry both their target URI and the visible label.
paragraph.Hyperlinks.Add(new Hyperlink
{
	ExternalUri = new Uri("https://example.com/project-details", UriKind.Absolute),
	Display = "Project details"
});

document.Save(ExampleSupport.OutputPath("BookmarksAndHyperlinks.docx"));
