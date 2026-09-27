using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("مرحبا بالعالم");

// Bidi formatting marks the run as right-to-left text for Word's shaping and layout.
paragraph.Runs[0].Font.Bidi = true;
paragraph.Runs[0].Font.LocaleIdBi = 1025;

document.Save(ExampleSupport.OutputPath("RightToLeftText.docx"));
