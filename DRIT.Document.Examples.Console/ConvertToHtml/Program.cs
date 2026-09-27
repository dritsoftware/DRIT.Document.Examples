using DRIT.Document;
using DRIT.Document.Examples.Shared;
using DRIT.Document.Export.Html;

var document = new Document();
document.AddSection().AddParagraph("HTML export");
document.Sections[0].Body.AddParagraph("The document model can be exported to HTML without Word.");

// SaveAsHtml writes the document model as an HTML file.
document.SaveAsHtml(ExampleSupport.OutputPath("ConvertToHtml.html"));
