using DRIT.Document.Examples.Shared;
using DRIT.Document;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("Normal text: ");

// Format each run independently after appending it to the paragraph.
var bold = paragraph.AppendRun("bold");
bold.Font.Bold = true;
var italic = paragraph.AppendRun(", italic");
italic.Font.Italic = true;
italic.Font.Size = 16;

// Save the in-memory document as a Word file.
document.Save(ExampleSupport.OutputPath("CharacterFormatting.docx"));
