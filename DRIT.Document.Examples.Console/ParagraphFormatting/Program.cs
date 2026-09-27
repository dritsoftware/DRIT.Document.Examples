using DRIT.Document.Examples.Shared;
using DRIT.Document;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("Centered report introduction");

// ParagraphFormat controls alignment, indentation, and spacing independently of run formatting.
paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
paragraph.ParagraphFormat.LeftIndent = 24;
paragraph.ParagraphFormat.SpaceAfter = 18;

document.Save(ExampleSupport.OutputPath("ParagraphFormatting.docx"));
