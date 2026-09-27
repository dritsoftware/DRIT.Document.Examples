using DRIT.Document;
using DRIT.Document.Drawing;
using DRIT.Document.Examples.Shared;

var document = new Document();
var image = new Shape(document, ShapeType.Rectangle, 4, 3);

// SetImage replaces the shape's placeholder geometry with image content.
image.SetImage(ExampleSupport.EnsureOnePixelPng());
document.AddSection().AddParagraph("Product artwork");
document.Sections[0].Body.AddParagraph().AppendShape(image);

document.Save(ExampleSupport.OutputPath("Pictures.docx"));
