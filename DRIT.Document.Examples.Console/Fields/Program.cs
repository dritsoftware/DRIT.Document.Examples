using DRIT.Document;
using DRIT.Document.Examples.Shared;

var document = new Document();
document.AddSection().AddParagraph("Current author: ");
var fieldParagraph = document.Sections[0].Body.AddParagraph();

// FieldBuilder creates the instruction and the visible result without hand-writing XML.
fieldParagraph.Fields.Add(new FieldBuilder(FieldType.Author)
	.WithResult("NovaEdge Technologies")
	.Build());

document.Save(ExampleSupport.OutputPath("Fields.docx"));
