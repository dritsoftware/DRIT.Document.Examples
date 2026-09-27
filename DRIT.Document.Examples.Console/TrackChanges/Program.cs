using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();

// Enable revision tracking before adding content that Word should monitor.
document.DocumentSettings.TrackRevisions = true;
document.AddSection().AddParagraph("Tracked content");
document.Sections[0].Body.AddParagraph("Changes can be accepted or rejected through Revisions.");

document.Save(ExampleSupport.OutputPath("TrackChanges.docx"));
