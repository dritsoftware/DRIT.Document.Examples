using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Document settings");

// Settings control how Word opens and lays out the saved document.
document.DocumentSettings.Zoom = 110;
document.DocumentSettings.AutoHyphenation = true;
document.DocumentSettings.TrackRevisions = false;

document.Save(ExampleSupport.OutputPath("DocumentSettings.docx"));
