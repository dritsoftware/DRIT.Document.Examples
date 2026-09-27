using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("ODT export");
document.Sections[0].Body.AddParagraph("Cross-format output from the same document model.");

// The output format is selected from the .odt extension.
document.Save(ExampleSupport.OutputPath("ConvertToOdt.odt"));
