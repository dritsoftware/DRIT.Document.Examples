using DRIT.Document.Examples.Shared;
using DRIT.Document;

var document = new Document();
document.AddSection().AddParagraph("Landscape report");
var pageSetup = document.Sections[0].PageSetup;

// Page setup values are expressed in points and apply to the selected section.
pageSetup.Orientation = Orientation.Landscape;
pageSetup.LeftMargin = 50;
pageSetup.RightMargin = 50;

document.Save(ExampleSupport.OutputPath("PageSetup.docx"));
