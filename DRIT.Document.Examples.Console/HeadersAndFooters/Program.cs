using DRIT.Document.Examples.Shared;
using DRIT.Document;

var document = new Document();
var section = document.AddSection();

// Headers and footers belong to a section and repeat on its pages.
section.AddParagraph("Report body");
section.AddHeaderFooter(HeaderFooterType.HeaderPrimary).AddParagraph("NovaEdge Technologies");
section.AddHeaderFooter(HeaderFooterType.FooterPrimary).AddParagraph("Internal use only");

document.Save(ExampleSupport.OutputPath("HeadersAndFooters.docx"));
