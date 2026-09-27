using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Document properties");

// Built-in properties become package metadata that Word displays in its Info panel.
document.BuiltInDocumentProperties.Title = "Annual Research Report";
document.BuiltInDocumentProperties.Author = "NovaEdge Technologies";
document.BuiltInDocumentProperties.Subject = "Example metadata";
document.Save(ExampleSupport.OutputPath("DocumentProperties.docx"));
