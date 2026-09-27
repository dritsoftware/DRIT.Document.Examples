using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Notes use the section page setup.");
document.AddFootnote("Footnote with configured numbering.");

// Footnote and endnote numbering is configured on the section page setup.
document.Sections[0].PageSetup.FootnoteOptions.NumberStyle = NumberStyle.LowercaseLetter;
document.Sections[0].PageSetup.EndnoteOptions.NumberStyle = NumberStyle.UppercaseRoman;

document.Save(ExampleSupport.OutputPath("FootnoteOptions.docx"));
