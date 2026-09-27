using System;
using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("Review this paragraph before publishing.");

// Comments are document-level objects; their text and author are stored with the package.
var comment = new Comment
{
	Author = "NovaEdge reviewer",
	Initials = "NE",
	DateTime = DateTime.UtcNow,
	Text = "Please confirm the figures in this paragraph."
};
document.Comments.Add(comment);
paragraph.AddCommentAnchor(comment.Id);

document.Save(ExampleSupport.OutputPath("Comments.docx"));
