using DRIT.Document;
using DRIT.Document.Drawing;
using DRIT.Document.Examples.Shared;

var document = new Document();

// TextBoxShape combines drawing geometry with a text payload.
var textBox = new TextBoxShape(document, ShapeType.Rectangle, 8, 3)
{
	Text = "Quarterly status: on track"
};
document.AddSection().AddParagraph("Text box example");
document.Sections[0].Body.AddParagraph().AppendShape(textBox);

document.Save(ExampleSupport.OutputPath("TextBoxes.docx"));
