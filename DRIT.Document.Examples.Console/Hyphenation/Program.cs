using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Automatic hyphenation helps Word fit long words within narrow columns.");

// These settings are written to settings.xml and take effect when Word lays out the file.
document.DocumentSettings.AutoHyphenation = true;
document.DocumentSettings.HyphenationZone = 360;
document.DocumentSettings.ConsecutiveHyphenLimit = 3;

document.Save(ExampleSupport.OutputPath("Hyphenation.docx"));
