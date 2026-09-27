using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("View settings");

// ViewOptions controls the initial Word view and zoom when the file opens.
document.ViewOptions.ViewType = ViewType.Print;
document.ViewOptions.ZoomPercent = 125;

document.Save(ExampleSupport.OutputPath("ViewOptions.docx"));
