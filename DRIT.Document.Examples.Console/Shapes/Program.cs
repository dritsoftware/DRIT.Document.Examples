using DRIT.Document;
using DRIT.Document.Drawing;
using DRIT.Document.Examples.Shared;

var document = new Document();

// Create a shape in centimetres, configure it, and append it as an inline drawing.
var shape = new Shape(document, ShapeType.Rectangle, 5, 3);
shape.Name = "Report callout";
shape.RotationAngle = 12;
document.AddSection().AddParagraph("Drawing shape example");
document.Sections[0].Body.AddParagraph().AppendShape(shape);

document.Save(ExampleSupport.OutputPath("Shapes.docx"));
