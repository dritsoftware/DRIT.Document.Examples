using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
var paragraph = document.AddSection().AddParagraph("Equation: ");

// AppendOfficeMath stores the equation as an Office Math object in the paragraph.
paragraph.AppendOfficeMath("x^2 + y^2 = z^2", OfficeMathDisplayType.Display);

document.Save(ExampleSupport.OutputPath("OfficeMath.docx"));
