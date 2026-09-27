using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("A statement with notes.");

// Add each note to the document; the library assigns the next available note id.
var footnote = document.AddFootnote("Footnote detail");
var endnote = document.AddEndnote("Endnote detail");
paragraph.AppendFootnoteReference(footnote.Id);
paragraph.AppendEndnoteReference(endnote.Id);

document.Save(ExampleSupport.OutputPath("FootnotesAndEndnotes.docx"));
